import pyautogui as pag
import os
import pandas as pd
import pyperclip
import sys
import time
from win10toast import ToastNotifier

MENU_PRINCIPAL = (151, 245)
CONSULTA_DE_PROCESSOS = (178, 412)
BTN_PESQUISAR = (513, 662)
BTN_ANEXOS = (272, 689)
BTN_ABRIR_ARQUIVO = (1307, 966)
JANELA_CGPJ = (381, 21)
VOLTAR = (1247, 438)

BTN_ABRIR_ARQUIVO = (1315, 962)

msg = ToastNotifier()

cancelar_loop = False

# Manipulação de Excel
df = pd.read_excel('banco_de_avalista.xlsx', sheet_name='Sheet')
df = df.sort_values(by='GCPJ', ascending=True)
total_de_itens = df.shape[0]
linha = 0

def scroll():
    time.sleep(1)

    for _ in range(3):
        pag.press('pagedown')
    time.sleep(.5)
    try:
        anexo = pag.locateOnScreen('assets/anexo.png', confidence=.7)
        pag.click(anexo)
        time.sleep(.5)
    except Exception as e:
        print(f'Não foi possível localixar a aba de anexos.\n{e}')

    for _ in range(3):
        pag.press('pageup')

def jumper():
    doubleArrow = pag.locateOnScreen("assets/double_arrow.png", confidence=.7)
    time.sleep(.5)
    print('Pulando para a última página')
    pag.click(doubleArrow)
    time.sleep(1)

def abrir_arquivo():
    visualizar_arquivo = pag.locateOnScreen("assets/visualizar arquivos.png", confidence=.7)
    pag.click(visualizar_arquivo, duration=.5)
    time.sleep(.7)
    pag.click(BTN_ABRIR_ARQUIVO, duration=.5)
    pag.click(VOLTAR)
    pag.moveTo(4065, 911, duration= 1)

def voltarUm():
    back_arrow = pag.locateOnScreen("assets/back_arrow.png", confidence=.7)
    pag.click(back_arrow)

def localizarIP():
    try:
        ip = pag.locateOnScreen("assets/IP.png", confidence=.8)
        print("Arquivo IP econtrado!!\nAbrindo documento.")
        pag.moveTo(ip)
        pag.moveRel(-55,0)
        pag.click()
        abrir_arquivo()            
        return True

    except:
        print("Não foi possível localizar o IP")
        return voltarUm()

def processar_linha(valor_gcpj):
    pag.click(MENU_PRINCIPAL)
    pag.click(CONSULTA_DE_PROCESSOS, duration=.5)
    time.sleep(1.5)


    print(f'\nProcessando {int(valor_gcpj)} ...\n')
    
    num_processo_str = str(valor_gcpj)
    pyperclip.copy(num_processo_str)
    pag.hotkey('ctrl','v')
    time.sleep(.5)
    
    pag.click(BTN_PESQUISAR)
    time.sleep(1.5)
    scroll()

    time.sleep(1.5)
    doubleArrow = pag.locateOnScreen("assets/double_arrow.png", confidence=.8)
    print('Pulando para a última página')
    pag.click(doubleArrow)

    time.sleep(1.5)

    while True:
        ip_encontrado = localizarIP()
        time.sleep(1)
        if ip_encontrado == True:
            break
        if ip_encontrado == False:
            voltarUm()
    
    print('Processo finalizado')
    return

if __name__ == "__main__":
    try:    
        for processo in range(total_de_itens):
            gcpj = df.at[linha, 'GCPJ']
            input('\nIniciar automação...')
            processar_linha(gcpj)
            linha = linha + 1
    
    except Exception as e:
        print(f'Erro encontrado. Automação interrompiada.\n{e}')