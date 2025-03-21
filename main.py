import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import zipfile
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

def selecionar_arquivos():
    """Abre uma janela para selecionar um ou mais arquivos XML ou compactados"""
    arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos XML e Compactados", "*.xml;*.zip;*.rar")])
    if arquivos:
        entrada_arquivo.delete(0, tk.END)
        entrada_arquivo.insert(0, "; ".join(arquivos))

def extrair_arquivos_compactados(arquivos):
    """Extrai arquivos XML de arquivos compactados"""
    xml_files = []
    for arquivo in arquivos:
        if arquivo.endswith(".xml"):
            xml_files.append(arquivo)
        elif arquivo.endswith(".zip"):
            with zipfile.ZipFile(arquivo, 'r') as zip_ref:
                extracao_dir = os.path.dirname(arquivo)
                zip_ref.extractall(extracao_dir)
                for nome in zip_ref.namelist():
                    if nome.endswith(".xml"):
                        xml_files.append(os.path.join(extracao_dir, nome))
    return xml_files

def processar_xml(arquivo_xml):
    """Processa um arquivo XML e retorna os dados relevantes"""
    try:
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()
        
        # Extração dos dados do emissor
        emissor = root.find(".//{http://www.portalfiscal.inf.br/nfe}emit")
        if emissor is not None:
            emissor_info = {
                "CNPJ": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}CNPJ").text,
                "Razao Social": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}xNome").text,
                "Nome Fantasia": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}xFant").text,
                "Endereco": f"{emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xLgr').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}nro').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xBairro').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xMun').text} - {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}UF').text}"
            }
        else:
            emissor_info = {}
        
        # Extração da data de emissão
        ide = root.find(".//{http://www.portalfiscal.inf.br/nfe}ide")
        if ide is not None:
            dh_emi = ide.find(".//{http://www.portalfiscal.inf.br/nfe}dhEmi").text
            data_emissao = datetime.strptime(dh_emi, "%Y-%m-%dT%H:%M:%S-03:00").strftime("%d/%m/%Y")
        else:
            data_emissao = "Data não disponível"
        
        # Extração dos dados das notas
        dados_vendas = []
        notas = root.findall(".//{http://www.portalfiscal.inf.br/nfe}det")
        for nota in notas:
            valor = float(nota.find(".//{http://www.portalfiscal.inf.br/nfe}vProd").text)
            descricao = nota.find(".//{http://www.portalfiscal.inf.br/nfe}xProd").text
            dados_vendas.append([descricao, f"R$ {valor:.2f}", data_emissao])  # Adiciona a data ao lado do valor
        
        return emissor_info, dados_vendas, data_emissao
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo {arquivo_xml}: {e}")
        return {}, [], "Data não disponível"

def gerar_pdf():
    """Lê os XMLs e gera um relatório em PDF formatado"""
    arquivos_selecionados = entrada_arquivo.get().split("; ")
    xml_files = extrair_arquivos_compactados(arquivos_selecionados)
    
    if not xml_files:
        messagebox.showerror("Erro", "Nenhum arquivo XML encontrado")
        return
    
    total_vendas = 0
    dados_vendas = []
    emissor_info = {}
    data_geral = None  # Data geral do relatório (mês de referência)
    
    # Barra de progresso
    progresso = ttk.Progressbar(janela, orient="horizontal", length=300, mode="determinate")
    progresso.pack(pady=10)
    progresso["maximum"] = len(xml_files)
    
    try:
        with ThreadPoolExecutor() as executor:
            futures = {executor.submit(processar_xml, arquivo_xml): arquivo_xml for arquivo_xml in xml_files}
            for i, future in enumerate(as_completed(futures)):
                progresso["value"] = i + 1
                janela.update_idletasks()
                emissor_info_partial, dados_vendas_partial, data_emissao = future.result()
                if not emissor_info and emissor_info_partial:
                    emissor_info = emissor_info_partial
                dados_vendas.extend(dados_vendas_partial)
                total_vendas += sum(float(dado[1].replace("R$ ", "")) for dado in dados_vendas_partial)
                
                # Define a data geral do relatório (mês de referência)
                if data_geral is None:
                    data_geral = datetime.strptime(data_emissao, "%d/%m/%Y").strftime("%m/%Y")
        
        # Escolher local para salvar
        nome_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Arquivo PDF", "*.pdf")])
        if not nome_pdf:
            return
        
        # Criar PDF formatado
        doc = SimpleDocTemplate(nome_pdf, pagesize=A4)
        elementos = []
        styles = getSampleStyleSheet()
        
        elementos.append(Paragraph("<b>Relatório Fiscal</b>", styles['Title']))
        elementos.append(Spacer(1, 12))
        
        if emissor_info:
            elementos.append(Paragraph(f"<b>CNPJ:</b> {emissor_info['CNPJ']}", styles['Normal']))
            elementos.append(Paragraph(f"<b>Razão Social:</b> {emissor_info['Razao Social']}", styles['Normal']))
            elementos.append(Paragraph(f"<b>Nome Fantasia:</b> {emissor_info['Nome Fantasia']}", styles['Normal']))
            elementos.append(Paragraph(f"<b>Endereço:</b> {emissor_info['Endereco']}", styles['Normal']))
            elementos.append(Spacer(1, 12))
        
        # Adiciona a data geral do relatório
        if data_geral:
            elementos.append(Paragraph(f"<b>Mês de Referência:</b> {data_geral}", styles['Normal']))
            elementos.append(Spacer(1, 12))
        
        # Tabela com os dados das vendas (incluindo a data)
        tabela = Table([['Produto', 'Valor', 'Data']] + dados_vendas, colWidths=[250, 100, 100])
        tabela.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elementos.append(tabela)
        elementos.append(Spacer(1, 12))
        elementos.append(Paragraph(f"<b>Total de Vendas:</b> R$ {total_vendas:.2f}", styles['Normal']))
        
        doc.build(elementos)
        
        messagebox.showinfo("Sucesso", f"Relatório gerado: {nome_pdf}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar XML: {e}")
    finally:
        progresso.pack_forget()

# Criar a interface gráfica
janela = tk.Tk()
janela.title("Gerador de Relatório Fiscal")

# Campo para selecionar os arquivos
frame = tk.Frame(janela)
frame.pack(pady=20)
entrada_arquivo = tk.Entry(frame, width=50)
entrada_arquivo.pack(side=tk.LEFT, padx=5)
botao_selecionar = tk.Button(frame, text="Selecionar Arquivos", command=selecionar_arquivos)
botao_selecionar.pack(side=tk.LEFT)

# Botão para gerar o PDF
botao_gerar = tk.Button(janela, text="Gerar PDF", command=gerar_pdf)
botao_gerar.pack(pady=10)

janela.mainloop()