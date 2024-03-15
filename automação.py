# -*- coding: utf-8 -*-
import openpyxl
import keyboard
import pyautogui
from time import sleep
from urllib.parse import quote

pyautogui.PAUSE = 1
planilha = openpyxl.load_workbook("planilha.xlsx")
planilha = planilha['Planilha1']
pyautogui.press('win')
pyautogui.write('Microsoft Edge')
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 't')
sleep(2)

for linha in planilha.iter_rows(min_row=2):
    nome = linha[0].value
    mensagem = linha[2].value
    telefone = linha[1].value

    mensagem = f'Olá {nome}, estou apenas testando meu código, então ignore(teste2)'
    url = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    keyboard.write(url)
    sleep(1)
    pyautogui.press('enter')
    sleep(15)
    posicao = pyautogui.locateOnScreen('Capturar.PNG')
    sleep(3)
    if posicao is not None:
        # Se a imagem for encontrada, clique nela
        pyautogui.click(posicao)
        sleep(1)
    else:
        posicao = pyautogui.locateOnScreen('Capturar.PNG')
        sleep(1)
        pyautogui.click()
    sleep(1)
    pyautogui.click(x=755, y=50)
    sleep(1)
