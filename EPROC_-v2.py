from re import A
from time import sleep
from pandas.core.indexes.base import Index
from selenium.webdriver.common.keys import Keys
import xlrd
from selenium import webdriver
import pandas as pd
from pandas.io.parsers import TextParser
from lxml.html import parse
from urllib.request import urlopen
from easygui import*
import ctypes
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from urllib import request
import shutil
import unidecode
import os

#DADOS

login_espaider = 'teste'
senha_espaider = 'Sumerios0394&'


loginFelipe = 'RS111059'
senhaFelipe = 'Alegrianoesforço!'


loginMauricio = 'RS036798'
senhaMauricio = 'Barbieri5515*'

message = "Coletar número de precatório?"
title = "Coleta precatório"
if boolbox(message, title, ["Sim", "Não"]):
    coleta_prec = 'sim'
else:
    coleta_prec ='nao'

message = "Coletar nota de expediente?"
title = "Coleta nota"
if boolbox(message, title, ["Sim", "Não"]):
    coleta_nota = 'sim'
else:
    coleta_nota ='nao'

message = "Qual grau de jurísdição?"
title = "Jurisdição"
if boolbox(message, title, ["1º Grau", "2º Grau"]):
    url = 'https://eproc1g.tjrs.jus.br/eproc/'
else:
    url ='https://eproc2g.tjrs.jus.br/eproc/'

message = "Qual usuário?"
title = "Usuário"
if boolbox(message, title, ["Felipe", "Maurício"]):
    login = loginFelipe
    senha = senhaFelipe
    socio = 'nao'
else:
    login = loginMauricio
    senha = senhaMauricio
    socio = 'sim'



options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
"download.prompt_for_download": False, #To auto download the file
"plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
})
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
espaider = 'http://barbieriadvogados.dyndns.org:40400/Barbieri/'


#COLETA DADOS EXCEL
wb = xlrd.open_workbook('INTIMACOES TESTE.xls')
planilha = wb.sheet_by_name('Planilha1')
total_linhas = planilha.nrows
total_colunas = planilha.ncols

navegador.get(url)      #ABRE NAVEGADOR
sleep(2)

navegador.find_element_by_xpath('//*[@id="txtUsuario"]').send_keys(login)       #CAMPO LOGIN
sleep(0.5)

navegador.find_element_by_xpath('//*[@id="pwdSenha"]').send_keys(senha)     #CAMPO SENHA
sleep(0.5)

navegador.find_element_by_xpath('//*[@id="sbmEntrar"]').click()     #BOTAO ENTRAR

msgbox("Resolva o Captcha caso apareça!")

navegador.find_element_by_xpath('//*[@id="tr0"]').click()       #SELECIONA PERFIL ADVOGADO
sleep(1.5)

dados = []
precatorios =[]
precatorios2 = []
processos = []
procsadm = []
data_dist = []
comarcas = []

titulos =[]
tipos = []
requerentes = []
requeridos = []
notas = []
textoVermelho = []
cj_descricao = []
cj_cpf=[]
cj_linksPDF = []

#PESQUISA PROCESSOS PARA CADA LINHA DO EXCEL
for i in range(2, total_linhas):
  processo = planilha.cell_value(rowx=i, colx=0)
  processo = str(processo)
  print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
  print(processo)
  
  navegador.find_element_by_xpath('//*[@id="navbar"]/div/div[3]/div[4]/form/input[1]').send_keys(processo)      #ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
  sleep(0.5)

  navegador.find_element_by_xpath('//*[@id="navbar"]/div/div[3]/div[4]/form/button[1]').click()     #BOTAO PARA PESQUISAR
  sleep(1)
  
  try:
    processo_originario = navegador.find_element_by_xpath('//*[@id="tableRelacionado"]/tbody/tr/td[1]/font/a').text     #COLETA NÚMERO DO PROCESSO ORIGINÁRIO
    sleep(0.5)
  except NoSuchElementException:
    processo_originario = 'Sem dados de processo originário'
    
  data = navegador.find_element_by_xpath('//*[@id="txtAutuacao"]').text   #COLETA DATA DE DISTRIBUIÇÃO
  sleep(0.5)
  data_reduc = data[0:10]
  nomeParteTodos = navegador.find_elements_by_class_name('infraNomeParte')  #COLETA NOME DA PARTE
  nomeParte = nomeParteTodos[0].text
  sleep(0.5)
  print (nomeParte)
  cpf=navegador.find_element_by_xpath('//*[@id="spnCpfParteAutor0"]').text    #COLETA CPF DO AUTOR
  print(cpf)

  
  if coleta_prec == 'sim':
    try:

      precatorio = navegador.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[1]/div[1]/form[2]/div[2]/div[1]/div/fieldset[1]/div/table/tbody/tr[2]/td[1]/font/a').text   #COLETA NÚMERO PRECATÓRIO
      sleep(0.5)
      precatorio2 = navegador.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[1]/div[1]/form[2]/div[2]/div[1]/div/fieldset[1]/div/table/tbody/tr[1]/td[1]/font/a').text
      sleep(0.5)
      controlePrecatorio = 'Sim'
    except:
      precatorio = 'Sem número de precatório'
      controlePrecatorio = 'Não'
    precatorios.append(precatorio)    #JOGA O VALOR DO PRECATÓRIO PARA A LISTA
    precatorios2.append(precatorio2)    #JOGA O VALOR DO PROCESSO ORIGINÁRIO PARA A LISTA, CASO ALI ESTEJA O NUMERO DO PRECATORIO

  if coleta_nota == 'sim':
    controle_integra = navegador.find_element_by_xpath('//*[@id="fldAcoes"]/center/a[1]').text
    if controle_integra == "Acesso íntegra do processo":
      navegador.find_element_by_xpath('//*[@id="fldAcoes"]/center/a[1]').click()  #CLICA EM ACESSO A ÍNTEGRA
      sleep(1.5)
      navegador.switch_to_alert().accept()                                        #ACEITA O PRIMEIRO ALERTA
      sleep(4)
      try:
        WebDriverWait(navegador,6).until(EC.alert_is_present())
        alert = navegador.switch_to.alert
        alert.accept()
      except TimeoutException:
        print('Erro no segundo alerta a respeito do acesso à integra')
    sleep(1)
    controleClara = len(navegador.find_elements_by_css_selector('[class="infraTrClara infraEventoPrazoAguardando"]'))   #CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE CLARA
    controleEscura = len(navegador.find_elements_by_css_selector('[class="infraTrEscura infraEventoPrazoAguardando"]')) #CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE ESCURA

    controleClaraAmarelo = len(navegador.find_elements_by_css_selector('[class="infraTrClara infraEventoPrazoAberto"]'))   #CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE CLARA
    controleEscuraAmarelo = len(navegador.find_elements_by_css_selector('[class="infraTrEscura infraEventoPrazoAberto"]')) #CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE ESCURA

    if controleClaraAmarelo>0 or controleEscuraAmarelo>0:
      print('Prazo em aberto')
      processo_originario='Prazo em aberto'     #JOGA O VALOR EXTRAIDO PARA A LISTA
      processo='Prazo em aberto'     #JOGA O VALOR DO PROCESSO PARA A LISTA
      data_reduc='Prazo em aberto'    #JOGA VALOR DATA PARA A LISTA
      precatorio_cert='Prazo em aberto'
      tipo_acao='Prazo em aberto'
      requerente='Prazo em aberto'
      requerido='Prazo em aberto'
      paragrafoUnificado='Prazo em aberto'

    elif controleClara > 0 or controleEscura>0:
      if controleClara > 0:
        print ("Prazo Aguardando Abertura - linha clara")
        vermelho = navegador.find_elements_by_css_selector('[class="infraTrClara infraEventoPrazoAguardando"]')
      elif controleEscura > 0:
        print ("Prazo Aguardando Abertura - linha escura")    
        vermelho = navegador.find_elements_by_css_selector('[class="infraTrEscura infraEventoPrazoAguardando"]')
      for x in range (0 , len(vermelho)):
        textoVermelho.append(vermelho[x].text)
        if nomeParte[0] in vermelho[x].text:
              controleTexto = x
              break
          
      textoVermelho[controleTexto] = textoVermelho[controleTexto].replace('\n','')        #TIRA AS QUEBRAS DE LINHA
      posterior=textoVermelho[controleTexto].split(':')                                   #ARMAZENAMENTO DO PRIMEITO TERMO ATÉ O PRÓXIMO
      anterior=textoVermelho[controleTexto].split('(')                                    #menos um do posterior
      indiceParenteses = posterior[3].find('(')
      resultado1= posterior[3][0:indiceParenteses]                                               #UNIÃO DOS RESULTADOS
      resultado2=resultado1.split(' ')

      
      if len(resultado1)>2:
        resultado=resultado2[1]
      else:
        resultado = resultado1.replace(' ','')
      print('RESULTADO: \n',resultado)
      navegador.find_element_by_css_selector(f'[id="trEvento{resultado}"]')
      descricao = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[3]/label').text
      cj_descricao.append(descricao)
      #um evento só
      descricaoNota = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[3]/label').text
      print('Descrição da Nota: \n', descricaoNota)
      try:
        tipoArquivo = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[5]/a').get_attribute('data-mimetype') 
        nomePDF = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[5]/a').text
        navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[5]/a').click()    #ABRE O ÚLTIMO DOCUMENTO(EVENTO)
        sleep(1)
        navegador.switch_to_window(navegador.window_handles[1])   #TROCA DE GUIA
      except:
        tipoArquivo = 'Sem dados'
      
      print(tipoArquivo)    

      if tipoArquivo == 'pdf':
        iframe = navegador.find_element_by_xpath('//*[@id="conteudoIframe"]')     # Pega o XPath do iframe e atribui a uma variável
        navegador.switch_to.frame(iframe)       # Muda o foco para o iframe
        navegador.find_element_by_xpath('//*[@id="open-button"]').click()
        sleep(2)
        navegador.switch_to.default_content() 
        nomePDFCerto = unidecode.unidecode(nomePDF)
        nomeArquivo = str(resultado) + '_' + nomePDFCerto
        print(nomeArquivo)
        os.mkdir(rf'C:\Users\Vitor Moreira\Desktop\python\EPROC Notas\{processo}')
        shutil.move(rf'C:\Users\Vitor Moreira\Downloads\{nomeArquivo}.pdf', rf'C:\Users\Vitor Moreira\Desktop\python\EPROC Notas\{processo}\{nomeArquivo}.pdf')
        precatorio_cert= 'Sem dados'
        tipo_acao= 'Sem dados'
        requerente= 'Sem dados'
        requerido= 'Sem dados'
        paragrafoUnificado = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[5]').text + '\n' + navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[3]').text
        linkPDF = rf'C:\Users\Vitor Moreira\Desktop\python\EPROC Notas\{nomeArquivo}.pdf'
        sleep(1)
        navegador.close()
        navegador.switch_to_window(navegador.window_handles[0])
      elif tipoArquivo == 'html': 
        linkPDF = 'Sem dados' 
        req2 = navegador.find_elements_by_class_name('nome_parte')                #COLETA TODOS OS REQUERIDOS
        prec = navegador.find_elements_by_class_name('identificacao_processo')    #COLETA O NUMERO DO PROCESSO
        try:
          precatorio_cert = prec[1].text
        except:
          precatorio_cert = prec[0].text                                          #PULA A CLASSE VAZIA E COLETA O DADO
          #precatorio_cert_texto = precatorio_cert[14:]                              #DIVIDE O TEXTO PRO DATA FRAME
        try:
          acao = navegador.find_element_by_class_name('assunto_processo').text    #COLETA A ACAO TODA
        except: 
          acao = 'Sem dados'
        requerente = req2[0].text                                                 #COLETA SÓ O REQUERENTE
        

        paragrafo = navegador.find_elements_by_class_name('paragrafoPadrao')
        contLinha=len(paragrafo)
        x=0
        paragrafoLinhas=[]
        for i in range (0,contLinha):
          linha=paragrafo[x].text
          x+=1
          paragrafoLinhas.append(linha)
        paragrafoUnificado="".join(paragrafoLinhas)

        requerido = req2[1].text                    #COLETA SÓ O REQUERENTE     
        if acao == 'Sem dados': 
          tipo_acao = 'Sem dados'
        else:                                                                                             
          tipo_acao = acao[14:]                                                     #DIVIDE O TIPO DA AÇÃO PARA O DATAFRAME
        sleep(1)
        navegador.close()
        navegador.switch_to_window(navegador.window_handles[0])
      else:
        paragrafoUnificado = navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[5]').text + '\n' + navegador.find_element_by_xpath(f'//*[@id="trEvento{resultado}"]/td[3]').text
    else:
      print('Prazo fechado ou erro')


    dados.append(processo_originario)     #JOGA O VALOR EXTRAIDO PARA A LISTA
    processos.append(processo)     #JOGA O VALOR DO PROCESSO PARA A LISTA
    data_dist.append(data_reduc)    #JOGA VALOR DATA PARA A LISTA
    titulos.append(precatorio_cert)
    tipos.append(tipo_acao)
    requerentes.append(requerente)
    requeridos.append(requerido)
    notas.append(paragrafoUnificado)
    cj_cpf.append(cpf)
    cj_linksPDF.append(linkPDF)
  else:
    requerente=navegador.find_element_by_xpath('//*[@id="tblPartesERepresentantes"]/tbody/tr[2]/td[1]/a[2]').text
    print('Erro geral')
 
    dados.append(processo_originario)     #JOGA O VALOR EXTRAIDO PARA A LISTA
    processos.append(processo)     #JOGA O VALOR DO PROCESSO PARA A LISTA
    data_dist.append(data_reduc)    #JOGA VALOR DATA PARA A LISTA
    requerentes.append(requerente)
    cj_cpf.append(cpf)

  textoVermelho = []
  print(f'Contagem: {i}/{total_linhas}')

if coleta_nota == 'sim':            
  df={'Principal':processos, 'Originário':dados, 'Data distribuição':data_dist, 'Título':titulos,'Tipo':tipos,'Requerente':requerentes,'CPF':cj_cpf,'Requerido':requeridos,'Certidão':notas, 'Link PDF':cj_linksPDF}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de notas.xlsx')        #CRIA O ARQUIVO EXCEL
elif coleta_prec =='sim':
  df={'Principal':processos, 'Originário':dados, 'Precatorios':precatorios, 'Precatorios 2':precatorios2, 'Data distribuição':data_dist,'Requerente':requerentes,'CPF':cj_cpf}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de processos com PRECATORIO.xlsx')        #CRIA O ARQUIVO EXCEL  

else:
  
  df={'Principal':processos, 'Originário':dados, 'Data distribuição':data_dist,'Requerente':requerentes,'CPF':cj_cpf}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de processos com CPF.xlsx')        #CRIA O ARQUIVO EXCEL

print('FINALIZADO')
navegador.quit()
    
'''
if coleta_prec == 'sim':
  for x in precatorios:
    link_prec = 'https://www.tjrs.jus.br/site_php/precatorios/precatorio.php?aba_opcao_consulta=numero&tipo_pesquisa=por_precatorio&btnPEsquisar=Pesquisar&Numero_Informado=' + x
    navegador.get(link_prec)
    sleep(5)
    try:
      processo_adm = navegador.find_element_by_xpath('//*[@id="conteudo"]/table[1]/tbody/tr[4]/td[2]').text   #COLETA O PROCESSO ADMINISTRATIVO
      sleep(0.5)
      comarca = navegador.find_element_by_xpath('//*[@id="conteudo"]/table[1]/tbody/tr[6]/td[2]').text    #COLETA COMARCA
      sleep(1)
    except:
      processo_adm = 'Erro'
      comarca = 'Erro'
    procsadm.append(processo_adm)
    comarcas.append(comarca)

#navegador.quit()

if coleta_prec == 'sim':
  df={'Principal':processos, 'Originário':dados, 'Data distribuição':data_dist,'Precatórios':precatorios, 'Processos ADM':procsadm, 'Comarca':comarcas, 'Título':titulos,'Requerente':requerentes,'Requerido':requeridos,'Certidão':notas}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de notas com precatórios.xlsx')        #CRIA O ARQUIVO EXCEL
elif coleta_nota == 'sim':
  df={'Principal':processos, 'Originário':dados, 'Data distribuição':data_dist, 'Título':titulos,'Requerente':requerentes,'Requerido':requeridos,'Certidão':notas}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de notas.xlsx')        #CRIA O ARQUIVO EXCEL
else:
  df={'Principal':processos, 'Originário':dados, 'Data distribuição':data_dist}      #ARMAZENA OS VALORES EM UMA MATRIZ
  df1=pd.DataFrame(df)        #CRIA O DATAFRAME
  df1.to_excel('./Lista de processos originários.xlsx')

navegador.quit()
ctypes.windll.user32.MessageBoxW(0, "Robo notas!", "Concluído!", 1)
#CONTROLE
print(df1)
'''