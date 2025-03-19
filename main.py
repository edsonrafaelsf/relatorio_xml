import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import A4 # type: ignore
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer # type: ignore
from reportlab.lib import colors # type: ignore
from reportlab.lib.styles import getSampleStyleSheet # type: ignore
import zipfile
import os

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
    
    try:
        for arquivo_xml in xml_files:
            tree = ET.parse(arquivo_xml)
            root = tree.getroot()
            
            # Extração dos dados do emissor
            if not emissor_info:
                emissor = root.find(".//{http://www.portalfiscal.inf.br/nfe}emit")
                if emissor is not None:
                    emissor_info = {
                        "CNPJ": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}CNPJ").text,
                        "Razao Social": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}xNome").text,
                        "Nome Fantasia": emissor.find(".//{http://www.portalfiscal.inf.br/nfe}xFant").text,
                        "Endereco": f"{emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xLgr').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}nro').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xBairro').text}, {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}xMun').text} - {emissor.find('.//{http://www.portalfiscal.inf.br/nfe}UF').text}"
                    }
            
            # Extração dos dados das notas
            notas = root.findall(".//{http://www.portalfiscal.inf.br/nfe}det")
            for nota in notas:
                valor = float(nota.find(".//{http://www.portalfiscal.inf.br/nfe}vProd").text)
                total_vendas += valor
                descricao = nota.find(".//{http://www.portalfiscal.inf.br/nfe}xProd").text
                dados_vendas.append([descricao, f"R$ {valor:.2f}"])
        
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
        
        tabela = Table([['Produto', 'Valor']] + dados_vendas, colWidths=[350, 100])
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
