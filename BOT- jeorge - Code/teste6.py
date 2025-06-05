import pandas as pd
from selenium import webdriver
import pyperclip
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import time
from selenium.webdriver.common.keys import Keys

# === CONFIGURAÇÕES ===
OUTLOOK_URL = "https://outlook.office365.com/mail/"
ARQUIVO_DESTINATARIOS = "destinatarios.xlsx"

df = pd.read_excel(ARQUIVO_DESTINATARIOS)

# === INICIA EDGE COM OPÇÕES ===
options = EdgeOptions()
options.add_argument('--start-maximized')
driver = webdriver.Edge(options=options)
driver.get(OUTLOOK_URL)

time.sleep(3)

# === AGUARDA LOGIN AUTOMATICAMENTE (espera aparecer botão "Novo email") ===
wait = WebDriverWait(driver, 120)  # aguarda até 2 minutos

time.sleep(2)

def clicar_novo_email(wait):
    for _ in range(3):  # tenta 3 vezes
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Novo email"]')))
            btn.click()
            return
        except StaleElementReferenceException:
            time.sleep(1)  # aguarda 1 seg para tentar de novo
    raise Exception("Falha ao clicar no botão Novo email após 3 tentativas")

time.sleep(1)

for index, row in df.iterrows():
    
    clicar_novo_email(wait)
    
    nome = row["nome"]
    email = row["email"]
    assunto = row["assunto"]
    corpo_original = row["corpo"]

    # Cria o corpo com saudação personalizada
    corpo_completo = f"Prezado(a) {nome},\n\n{corpo_original}"
    
    # Preenche o campo "Para"
    campo_para = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="textbox" and @aria-label="Para"]')))
    campo_para.click()
    campo_para.send_keys(email)
    time.sleep(1)
    campo_para.send_keys(Keys.ENTER)  # envia Enter após o email
    
    time.sleep(2)

    # Preenche o campo "Assunto"
    campo_assunto = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@aria-label="Assunto"]')))
    campo_assunto.click()
    time.sleep(2)
    campo_assunto.send_keys(assunto)
    time.sleep(2)

    # Copia o corpo personalizado para o clipboard
    pyperclip.copy(corpo_completo)

    # Cola no corpo do e-mail (elemento div com <br> e fonte Aptos)
    div_alvo = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[br and contains(@style, "font-family: Aptos")]')))
    div_alvo.click()
    time.sleep(1)
    div_alvo.send_keys(Keys.CONTROL, 'v')
    time.sleep(2)

    # Clica no botão "Enviar"
    btn_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Enviar"]')))
    btn_enviar.click()
    time.sleep(3)

    print(f"📤 E-mail enviado para: {nome} <{email}>")
    time.sleep(3)  # pequena pausa entre envios

# === FINALIZA ===
print("✅ Todos os e-mails foram enviados com sucesso!")
time.sleep(20)
driver.quit()
