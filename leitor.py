import pyautogui
import pyperclip
import time
import os
import PyPDF2
import re

# Função para copiar
def copy():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.1)  # Aguarde um curto período para a cópia ser concluída

# Função para colar
def paste(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')

# Função para extrair informações de um PDF
def extrair_informacoes(pdf_filename):
    pdf_file = open(pdf_filename, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    page = pdf_reader.pages[0]
    page_text = page.extract_text()

    expressao_titulo = r'TÍTULO\s+([\s\S]*?)TIPO\s+'
    expressao_tipo = r'TIPO\s+(.*?)\s+'
    expressao_sistema = r'SISTEMA\s+(.*?)\s+'
    expressao_linha = r'LINHA\s+(.*?)\s+'
    expressao_projetista = r'PROJETISTA\s+(.*?)\s+'
    expressao_trecho = r'TRECHO\s+(.*?)\s+'
    expressao_subtrecho = r'SUBTRECHO\s+(.*?)\s+'
    expressao_subsistema_conjunto = r'SUBSISTEMA/CONJUNTO\s+(.*?)\s+'
    expressao_area = r'ÁREA\s+(.*?)\s+'
    expressao_contrato = r'Nº CONTRATO\s+(.*?)\s+'
    expressao_etapa = r'ETAPA\s+(.*?)\s+'
    expressao_classe_subclasse = r'CLASSE/SUBCLASSE\s+(.*?)\s+'
    expressao_sequencial = r'SEQUENCIAL\s+(.*?)\s+'
    expressao_controle = r'Nº CONTROLE\s+(.*?)\s+'
    expressao_verificacao_data = r'VERIFICAÇÃO/DATA\s+(.*?)\s+'
    expressao_identificacao = r'IDENTIFICAÇÃO\s+(.*?)\s+'
    expressao_revisao = r'REVISÃO\s+(.*?)\s+'

    titulo_match = re.search(expressao_titulo, page_text)
    tipo_match = re.search(expressao_tipo, page_text)
    sistema_match = re.search(expressao_sistema, page_text)
    linha_match = re.search(expressao_linha, page_text)
    projetista_match = re.search(expressao_projetista, page_text)
    trecho_match = re.search(expressao_trecho, page_text)
    subtrecho_match = re.search(expressao_subtrecho, page_text)
    subsistema_conjunto_match = re.search(expressao_subsistema_conjunto, page_text)
    area_match = re.search(expressao_area, page_text)
    contrato_match = re.search(expressao_contrato, page_text)
    etapa_match = re.search(expressao_etapa, page_text)
    classe_subclasse_match = re.search(expressao_classe_subclasse, page_text)
    sequencial_match = re.search(expressao_sequencial, page_text)
    controle_match = re.search(expressao_controle, page_text)
    verificacao_data_match = re.search(expressao_verificacao_data, page_text)
    identificacao_match = re.search(expressao_identificacao, page_text)
    revisao_match = re.search(expressao_revisao, page_text)

    if titulo_match:
        print("TÍTULO:", titulo_match.group(1).strip())
    if tipo_match:
        print("TIPO:", tipo_match.group(1))
    if sistema_match:
        print("SISTEMA:", sistema_match.group(1))
    if linha_match:
        print("LINHA:", linha_match.group(1))
    if projetista_match:
        print("PROJETISTA:", projetista_match.group(1))
    if trecho_match:
        print("TRECHO:", trecho_match.group(1))
    if subtrecho_match:
        print("SUBTRECHO:", subtrecho_match.group(1))
    if subsistema_conjunto_match:
        print("SUBSISTEMA/CONJUNTO:", subsistema_conjunto_match.group(1))
    if area_match:
        print("ÁREA:", area_match.group(1))
    if contrato_match:
        print("Nº CONTRATO:", contrato_match.group(1))
    if etapa_match:
        print("ETAPA:", etapa_match.group(1))
    if classe_subclasse_match:
        print("CLASSE/SUBCLASSE:", classe_subclasse_match.group(1))
    if sequencial_match:
        print("SEQUENCIAL:", sequencial_match.group(1))
    if controle_match:
        print("Nº CONTROLE:", controle_match.group(1))
    if verificacao_data_match:
        print("VERIFICAÇÃO/DATA:", verificacao_data_match.group(1))
    if identificacao_match:
        print("IDENTIFICAÇÃO:", identificacao_match.group(1))
    if revisao_match:
        print("REVISÃO:", revisao_match.group(1))

# Caminho fixo da pasta onde os arquivos PDF estão localizados
diretorio_pdfs = "C:\\Users\\garot\\Desktop\\Teste final - Copia"

# Lista de arquivos na pasta
arquivos = os.listdir(diretorio_pdfs)

# Filtra apenas os arquivos PDF
arquivos_pdf = [arquivo for arquivo in arquivos if arquivo.lower().endswith(".pdf")]

# Solicita ao usuário quantos documentos serão cadastrados
nDocuments = int(pyautogui.prompt(text='Digite quantos documentos serão cadastrados', title='Cadastro Automático de documentos', default=''))

pyautogui.PAUSE = 0.5
time.sleep(1.0)

i = 0

for arquivo_pdf in arquivos_pdf[:nDocuments]:
    i = i + 1

    # Constrói o caminho completo do arquivo PDF
    pdf_filename = os.path.join(diretorio_pdfs, arquivo_pdf)

    # Extrair informações do arquivo PDF
    extrair_informacoes(pdf_filename)
