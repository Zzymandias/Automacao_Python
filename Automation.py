import pandas as pd
import pywhatkit as kit
from datetime import datetime
import time
import pyautogui

# Carregar a planilha
df = pd.read_excel("dados.xlsx")

hoje = datetime.now().strftime("%Y-%m-%d") 

print(f"Verificando mensagens para a data: {hoje}")

# Filtrar e enviar
for index, linha in df.iterrows():
    data_evento = str(linha['Data:']).split(' ')[0]
    
    if data_evento == hoje:
        nome = linha['Nome:']
        numero = str(linha['Numero:'])
        if not numero.startswith('+'):
            numero = "+" + numero
        
        assunto = linha['Assunto:']
        mensagem = f"Olá {nome}, estou passando para avisar sobre: {assunto}. Até logo!"
        
        print(f"Abrindo conversa com {nome} ({numero})...")
        
        kit.sendwhatmsg_instantly(numero, mensagem, wait_time=35, tab_close=False)
        
        time.sleep(2)
        
        pyautogui.press('enter')
        print(f"Mensagem enviada para {nome}!")
        
        # Espera o WhatsApp processar o envio antes de fechar a aba
        time.sleep(5)
        
        pyautogui.hotkey('ctrl', 'w')
        print("Aba fechada.")
        
        # Pausa maior entre pessoas diferentes para evitar bloqueio do WhatsApp
        time.sleep(10)

print("Processo concluído!")