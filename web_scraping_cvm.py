from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from selenium.webdriver.common.by import By
from selenium import webdriver
from time import sleep
import smtplib
import os

#Constantes
download_path = os.path.join(os.path.expanduser('~'),r'/Users/victormorais/Documentos/pasta_teste_downloads')
destinatario = "teste@teste.com.br"
assunto = "Teste de E-mail - Com Anexo de Exemplo"
corpo = "Olá isso é um teste! - Com Anexo de Exemplo"
caminho_pdf = r"/Users/victormorais/Documentos/pasta_teste_downloads"

#função para enviar um e-mail
def enviar_email(destinatario, assunto, corpo, caminho_pdf):
    # Configurações do remetente (seu email e senha de app)
    remetente = "teste@teste.com.br"
    senha_app = ""
    # Criação da mensagem
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    # Anexar o corpo do email
    msg.attach(MIMEText(corpo, 'plain'))
    # Função para anexar arquivos
    def anexar_arquivo(nome_arquivo, caminho_arquivo):
        # Ler o arquivo binário
        with open(caminho_arquivo, "rb") as arquivo:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(arquivo.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename={nome_arquivo}')
            msg.attach(parte)
    # Anexar o PDF
    anexar_arquivo("arquivo.pdf", caminho_pdf)
    # Configuração do servidor SMTP usando SMTP_SSL na porta 465
    try:
        servidor = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        servidor.login(remetente, senha_app)  # Fazer login com o email e senha de app
        # Enviar o email
        texto = msg.as_string()
        servidor.sendmail(remetente, destinatario, texto)
        # Encerrar a conexão
        servidor.quit()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Falha ao enviar email: {e}")



#inicia o chrome ignorando o erro
#print('INICIAR CHROME IGNORANDO CERTIFICADO INVALIDO')
options = webdriver.ChromeOptions()

options.add_argument('--ignore-ssl-errors=yes')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--disable-blink-features=AutomationControlled') # Pular Captcha
options.add_argument("--kiosk-printing")
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "savefile.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
    "safebrowsing.enabled": True,
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":"","capabilities":{}}],"selectedDestinationId":"Save as PDF","version":2}'
})
driver = webdriver.Chrome(options=options)
driver.get(url='https://cvmweb.cvm.gov.br/swb/default.asp?sg_sistema=fundosreg')
# driver.minimize_window()

cnpj = '09.215.250/0001-13'
#WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtCNPJNome"]')))
driver.switch_to.frame(1)
sleep(2)
driver.find_element(By.XPATH,'//*[@id="txtCNPJNome"]').click()
driver.find_element(By.XPATH,'//*[@id="txtCNPJNome"]').send_keys(cnpj)
sleep(2)
driver.find_element(By.XPATH,'//*[@id="btnContinuar"]').click()
sleep(2)
driver.find_element(By.XPATH,'//*[@id="ddlFundos__ctl0_Linkbutton2"]').click()
sleep(2)
driver.find_element(By.XPATH,'//*[@id="Hyperlink1"]/font').click()
sleep(2)
driver.execute_script("window.print()")
sleep(10)
enviar_email(destinatario, assunto, corpo, caminho_pdf)
driver.close()