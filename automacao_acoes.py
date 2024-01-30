"""
Projeto Python responsável por pegar todos os dados de ações no Yahoo Finance e enviar para um email 
pelo gmail.

Todo o processo é feito de forma automatica, basta inserir os nomes das ações que deseja ser analisadas
e o email para enviar.

OBS: É necessario estar com o login feito no gmail e ter o Google Chrome instalado no computador.
"""

import pyautogui as pa
import yfinance as yf
from time import sleep 
from pyperclip import copy
from datetime import datetime as dt


data_atual = dt.now().date().strftime("%d/%m/%Y")
nomes_acoes = [x for x in input("Digite os nomes dos tickers separados por espaço ").split(" ")]
email = input("Digite o email do destinatário ")
dados_acoes = []


for nome_acao in nomes_acoes:
    """Pegando os dados de cada ação passada"""
    acao = yf.Ticker(nome_acao).history(period="6mo")
    fechamento = acao.Close # Pegando a coluna de fechamento da ação

    maxima = fechamento.max()
    minima = fechamento.min()
    atual = fechamento.iloc[-1]
    
    dados_acoes.append({'nome': nome_acao.upper(), 'atual': atual, 'maxima': maxima, 'minima': minima})


# Automatizando o envio de email
pa.PAUSE = 1.5 

# Abrindo o navegador
pa.press('win')
pa.write('google chrome')
pa.press('enter')
sleep(1.5)


# Abrindo o gmail
pa.write('https://mail.google.com/mail/u/0/?pli=1#inbox')
pa.press('enter')
sleep(1.5)

for i in dados_acoes:
    # Criando um email e digitando o destinatário
    pa.click(x=131, y=267)
    copy(email)
    pa.hotkey('ctrl', 'v')
    pa.press('tab')

    # Colocando o assunto
    assunto = f"Fechamento da ação {i['nome']} do dia {data_atual}"
    copy(assunto)
    pa.hotkey('ctrl', 'v')
    pa.press('tab')
    
    # Corpo do email
    corpo = f"""
Prezado gestor

Segue as analises da ação {i['nome']}

Atual = {i['atual']}
Maxima = {i['maxima']}
Minima = {[i['minima']]}    

Qualquer dúvida estarei a disposição.
"""
    copy(corpo)
    pa.hotkey('ctrl', 'v')

    # Enviando o email
    pa.hotkey('ctrl','enter')
    sleep(1)


# Fechando o google
pa.hotkey('alt', 'f4')
