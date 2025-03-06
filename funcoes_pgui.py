import pyautogui as pgui
import time

kits = kits = ["430035", "396864", "702310"]

def tratamento(cod_item):
    #Verifica se o item é tratamento para unhas
    if cod_item in kits:
        pgui.moveTo(570,430)
        pgui.doubleClick()
        pgui.moveTo(570,470)
        pgui.doubleClick()
    else:
        pgui.moveTo(570,430)
        pgui.doubleClick()

def incompativel(produto_incompativel):
    if produto_incompativel == None:                
        time.sleep(1)
        pgui.press('right')
        pgui.press('enter')
    else:
        time.sleep(1)
        pgui.press('enter')
        time.sleep(1)
        pgui.press('right')
        pgui.press('enter')

def estoque(sem_estoque):
    if sem_estoque != None:
        pgui.write("MILIBOT")
        pgui.press('tab')
        pgui.press('tab')
        pgui.write("77777777")
        pgui.press('enter')

def cadastro(nome_cli, cpf_cli, telefone):    
    #clica para incluir cadastro
    pgui.moveTo(525,270)
    pgui.click()
    time.sleep(1)
    
    #clica e insere informação de nome
    pgui.moveTo(225,240)
    pgui.click()
    pgui.write(nome_cli)
    
    #clica e insere informação de CPF
    pgui.moveTo(85,320)
    pgui.click()
    pgui.write(cpf_cli)
    pgui.moveTo(775,320)
    pgui.click()
    
    #insere sexo
    pgui.write("M")
    pgui.press('tab')
    
    #insere telefone
    pgui.moveTo(200,580)
    pgui.click()
    pgui.write(telefone)
    pgui.hotkey('Alt', 'G')

def tela_inclusao(tela_incluir):
    if tela_incluir == None:
        #Abre o executar
        pgui.hotkey("win", "r")
        
        #Digita o caminho do smart
        pgui.write(r"\\10.1.1.5\j\smartpharmacy\SmartPhar.exe")
        
        #Realiza as interações para abrir o programa
        pgui.press("enter")
        time.sleep(2)
        pgui.press("left")
        pgui.press("enter")
        
        #Aguarda carregamento do programa, e insere as credenciais
        time.sleep(10)
        pgui.write("MILIBOT")
        pgui.press('TAB')
        pgui.press('TAB')
        pgui.write('77777777')
        pgui.press('TAB')
        pgui.write
        pgui.press('down')
        pgui.press('down')
        pgui.press('down')
        pgui.press('down')
        pgui.press('down')
        pgui.press('enter')
        time.sleep(2)
        pgui.press('enter')
        pgui.press('enter')
        time.sleep(2)
        pgui.press('enter')
        
        #Aguarda o smart abrir e abre as receitas
        receita = pgui.locateOnScreen("imagens/receita_smart.png", confidence=0.9)
        if receita != None:
            pgui.moveTo(20,85)
            pgui.click()
            
        #Clica no botão de Incluir Normal
        time.sleep(3)
        pgui.moveTo(450,130)
        pgui.click()