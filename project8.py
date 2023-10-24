import pyautogui
import pyperclip
import time
import os
import PyPDF2
import re

# Função para colar cada informação de atributo
def atributo(tipo):
    paste(tipo)
    time.sleep(0.1)
    pyautogui.press("enter")

# Função tab genérica
def tab():
    time.sleep(0.1)
    pyautogui.press("tab")
    time.sleep(0.1)

# Função do tab + atributos
def colarInformação(tipo):
    if tipo is not None:
        atributo(tipo)
        tab()

# Função para copiar
def copy():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.1)

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
    subsistema_conjunto_match = re.search(
        expressao_subsistema_conjunto, page_text)
    area_match = re.search(expressao_area, page_text)
    contrato_match = re.search(expressao_contrato, page_text)
    etapa_match = re.search(expressao_etapa, page_text)
    classe_subclasse_match = re.search(expressao_classe_subclasse, page_text)
    sequencial_match = re.search(expressao_sequencial, page_text)
    controle_match = re.search(expressao_controle, page_text)
    verificacao_data_match = re.search(expressao_verificacao_data, page_text)
    identificacao_match = re.search(expressao_identificacao, page_text)
    revisao_match = re.search(expressao_revisao, page_text)

    titulo = titulo_match.group(1).strip() if titulo_match else None
    tipo = tipo_match.group(1) if tipo_match else None
    sistema = sistema_match.group(1) if sistema_match else None
    linha = linha_match.group(1) if linha_match else None
    projetista = projetista_match.group(1) if projetista_match else None
    trecho = trecho_match.group(1) if trecho_match else None
    subtrecho = subtrecho_match.group(1) if subtrecho_match else None
    subsistema_conjunto = subsistema_conjunto_match.group(
        1) if subsistema_conjunto_match else None
    area = area_match.group(1) if area_match else None
    contrato = contrato_match.group(1) if contrato_match else None
    etapa = etapa_match.group(1) if etapa_match else None
    classe_subclasse = classe_subclasse_match.group(
        1) if classe_subclasse_match else None
    sequencial = sequencial_match.group(1) if sequencial_match else None
    controle = controle_match.group(1) if controle_match else None
    verificacao_data = verificacao_data_match.group(
        1) if verificacao_data_match else None
    identificacao = identificacao_match.group(
        1) if identificacao_match else None
    revisao = revisao_match.group(1) if revisao_match else None

    return titulo, tipo, sistema, linha, projetista, trecho, subtrecho, subsistema_conjunto, area, contrato, etapa, classe_subclasse, sequencial, controle, verificacao_data, identificacao, revisao

# Caminho fixo da pasta onde os arquivos PDF estão localizados
diretorio_pdfs = "C:\\Users\\garot\\Desktop\\Teste final - Copia"

# Lista de arquivos na pasta
arquivos = os.listdir(diretorio_pdfs)

# Filtra apenas os arquivos PDF
arquivos_pdf = [arquivo for arquivo in arquivos if arquivo.lower().endswith(".pdf")]

# Solicita ao usuário quantos documentos serão cadastrados
nDocuments = int(pyautogui.prompt(text='Digite quantos documentos serão cadastrados',
             title='Cadastro Automático de documentos', default=''))

pyautogui.PAUSE = 0.5
time.sleep(1.0)

# Primeiro Loop: Extrai informações dos documentos PDF
for i, arquivo_pdf in enumerate(arquivos_pdf[:nDocuments], start=1):
    # Constrói o caminho completo do arquivo PDF
    pdf_filename = os.path.join(diretorio_pdfs, arquivo_pdf)

    # Extrair informações do arquivo PDF
    titulo, tipo, sistema, linha, projetista, trecho, subtrecho, subsistema_conjunto, area, contrato, etapa, classe_subclasse, sequencial, controle, verificacao_data, identificacao, revisao = extrair_informacoes(
        pdf_filename)

    print(f"Documento {i}: ")
    if titulo:
        print("TÍTULO:", titulo)
    if tipo:
        print("TIPO:", tipo)
    if sistema:
        print("SISTEMA:", sistema)
    if linha == "7" or linha == "07":
        linha = "07 - Rubi - Luz - Jundiaí"
    if linha == "10":
        linha = "10 - Turquesa -"
    if linha == "11":
        linha = "11 - Coral - Barra Funda - Estudantes"
    if linha == "12":
        linha = "12 - Safira - Brás - Suzano"

    if linha:
        print("LINHA:", linha)
    if projetista:
        print("PROJETISTA:", projetista)
    if trecho:
        print("TRECHO:", trecho)
    if subtrecho:
        print("SUBTRECHO:", subtrecho)
    if subsistema_conjunto:
        print("SUBSISTEMA/CONJUNTO:", subsistema_conjunto)
    if area:
        print("ÁREA:", area)
    if contrato == "" or contrato == "ETAPA":
        contrato = "0"  # Define como "0" se for igual a NADA ou "ETAPA"
    if contrato:
        print("Nº CONTRATO:", contrato)
    if etapa:
        print("ETAPA:", etapa)
    if classe_subclasse:
        print("CLASSE/SUBCLASSE:", classe_subclasse)
    if sequencial:
        print("SEQUENCIAL:", sequencial)
    if controle:
        print("Nº CONTROLE:", controle)
    if verificacao_data:
        print("VERIFICAÇÃO/DATA:", verificacao_data)
    if identificacao:
        print("IDENTIFICAÇÃO:", identificacao)
    if revisao:
        print("REVISÃO:", revisao)
    print()

    # Inclua aqui o código para realizar as ações desejadas com as informações extraídas do PDF.

# Mudar de aba, colar o nome do n° de controle e pesquisar
pyautogui.hotkey('alt', 'tab')
time.sleep(1.5)

# Segundo Loop: Pesquisa e preenche informações em outro aplicativo
for i, arquivo_pdf in enumerate(arquivos_pdf[:nDocuments], start=1):
    # Inclua aqui o código para realizar a pesquisa com as informações desejadas.

    pdf_filename = os.path.join(diretorio_pdfs, arquivo_pdf)
    titulo, tipo, sistema, linha, projetista, trecho, subtrecho, subsistema_conjunto, area, contrato, etapa, classe_subclasse, sequencial, controle, verificacao_data, identificacao, revisao = extrair_informacoes(
        pdf_filename)

    controle = controle if controle else str(i)  # Atualize o controle

    paste(controle)

    time.sleep(0.2)
    pyautogui.press("enter")  # Pesquisa o documento
    time.sleep(7.50)  # Espera carregar
    pyautogui.press('tab', presses=29, interval=0.01)  # Clica no documento
    time.sleep(0.1)
    pyautogui.press("enter")  # Clica no documento

    time.sleep(15.00)  # Espera carregar
    pyautogui.press('tab', presses=2, interval=0.01)

    time.sleep(0.2)

    paste(titulo)  # Cola o Título

    pyautogui.press('tab', presses=23, interval=0.01)  # Clica na aba Atributos
    time.sleep(0.1)
    pyautogui.press("enter")  # Clica na aba Atributos
    time.sleep(0.3)
    pyautogui.press('tab', presses=2, interval=0.01)  # Chega na guia Sistema/Subsistema

    sistemaSubsistema = sistema + subsistema_conjunto  # Concatena Sistema/Subsistema

    # Cola o Sistema/Subsistema
    colarInformação(sistemaSubsistema)

    colarInformação(linha)

    linha = linha[:2] if linha else None  # Volta para apenas os 2 primeiros dígitos da linha

    # Concatena o Trecho/Subtrecho
    trechoSubtrecho = trecho + subtrecho

    colarInformação(trechoSubtrecho)

    colarInformação(etapa)

    colarInformação(classe_subclasse)

    classificação = tipo + sistema + linha + trechoSubtrecho + subsistema_conjunto + etapa + classe_subclasse + sequencial  # Concatena tudo para a identificação

    colarInformação(classificação)

    # Projeto (Por padrão "não identificado")
    naoIdentificado = ('Não identificado')
    colarInformação(naoIdentificado)

    colarInformação(projetista)

    colarInformação(contrato)

    tab()

    pyautogui.write("CPTM- CIA PTA TRENS METROP.")  # Coloca a empresa de origem

    # Fim da guia Atributos

    pyautogui.alert("ISSO É PARA OS TESTES!!!!!!! Confira se todas as informações dos atributos estão corretas")

    time.sleep(1.2)

    pyautogui.hotkey('alt', 'f4', interval=0.1)  # Volta para o Sesuite
    pyautogui.hotkey('f5')  # Reinicia o Sesuite
    time.sleep(5.5)
