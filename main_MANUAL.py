import pandas as pd
import xlrd as xlrd
import pyautogui as pgui
import tkinter as tk
import time
import requests
import pyperclip
import json
import cv2
import os
import numpy as np
import procurar_pasta
import combinar_arquivos
import ajuste_dataframe
import funcoes_pgui
import enviar_api
from datetime import datetime
from tkinter import filedialog
from tkinter import ttk
from tkcalendar import Calendar
from datetime import date
from procurar_pasta import proc_pasta

#Criação da interface gráfica no tkinter
root = tk.Tk()
root.title("Incluir Pedidos")
root.geometry("400x650")
root.resizable(0, 0)

canvas = tk.Canvas(root, width=200, height=200)
canvas.pack()

#Rótulo da tela
tk.Label(root, text="1º Selecione a pasta com os pedidos").place(x=50,y=30)

#Botão para procurar a pasta
button = tk.Button(root, text="Procurar", command=proc_pasta)
button.place(x=250, y=60)

#Função para juntar os arquivos da pasta em uma única planilha
def combine_excel_files():
    global separados
    
    #Armazena a variavel com o nome 
    folder_path = procurar_pasta.folder_path
    
    #Verifica se o caminho da pasta não está vazio
    if folder_path != None:
        combinar_arquivos.combina_arquivos()
    
    #Obtém o nome do arquivo
    output_file = combinar_arquivos.output_file
    
    #Consulta os pedidos Semi-Acabados no banco de dados
    response = requests.get("https://api.fmiligrama.com/produtos?tipoInterno=Semi-Acabado").json()
    
    #Armazena os itens Semi-Acabados em um dataframe
    semis = pd.json_normalize(response['data'])

    #Consulta os produtos industrializados no banco de dados
    response_industrializados = requests.get("https://api.fmiligrama.com/produtos?tipoInterno=Industrializado").json()

    #Armazena os itens em um dataframe
    industrializados = pd.json_normalize(response_industrializados['data'])

    #Juntar os produtos semi-acabados e industrializados
    semiacabados = pd.concat([semis, industrializados])

    print(semiacabados["codigo"])
    #Cria um Data Frame com a planilha de pedidos
    pedidos = pd.read_excel(folder_path + "/" + output_file, sheet_name='Sheet1')
    pedidos["Fone"].fillna("99999999999", inplace= True)
    pedidos["Celular"].fillna("99999999999", inplace= True)
    
    #Agrupa os itens duplicados
    pedidos = pedidos.groupby(["Número do pedido","Fone", "CPF/CNPJ", "Nome do contato", "Código (SKU)"]).sum("Quantidade")
    pedidos = pedidos.reset_index()
    
    #Converte a coluna codigo da base de dados para inteiro
    semiacabados = semiacabados[semiacabados['codigo'] != 'folder_sac']
    semiacabados["codigo"]=semiacabados["codigo"].astype(int)
    
    #Mescla a planilha de pedidos com a base de dados para filtrar os pedidos semi-acabados
    separados = pd.merge(pedidos, semiacabados, left_on="Código (SKU)", right_on="codigo", how="left")
    
    #Ajusta os dados do dataframe
    ajuste_dataframe.ajuste_excel(separados)
    print(separados)
    #Armazena os pedidos manipulados
    separados["tipoInterno"] = separados["tipoInterno"].str.strip()
    separados["tipoInterno"] = separados["tipoInterno"].str.lower()
    manipulados = separados[separados["tipoInterno"].isin(["manipulado"])]
    print(manipulados)
    #Soma a quantidade de reqs e pedidos
    total_pedidos = manipulados["Número do pedido"].drop_duplicates()
    total_pedidos = total_pedidos.count()
    total_requisicoes = manipulados["Quantidade"].sum()
    
    #Cria os rótulos para serem mostrados ao usuário
    total_pedidos_label.config(text=f"Total Pedidos: {total_pedidos}")
    total_requisicoes_label.config(text=f"Total REQs: {total_requisicoes}")

def incluir_smart():
    #Separa todos os itens que precisam ser inclusos no smart
    smart = separados["tipoInterno"] == "manipulado"
    smart = separados.loc[smart]
    
    #Variável para armazenar os números de pedidos
    num_pedido_ant = None
    num_pedido_ant = str(num_pedido_ant)
    reqs = []
    
    #Função para incluir os pedidos no smart
    for i in range(len(smart)):        
        #Define uma variável para cada valor a ser usado no código
        cod_item = smart["Código (SKU)"].iloc[i]
        nome_cli = smart["Nome do contato"].iloc[i]
        cpf_cli = smart["CPF/CNPJ"].iloc[i]
        num_ped = smart["Número do pedido"].iloc[i]
        qtd = smart["Quantidade"].iloc[i]
        telefone = smart["Fone"].iloc[i]
        observacoes = smart["Observações"].iloc[i]
        cod_item = str(cod_item)
        qtd = qtd = str(qtd)
        num_ped = str(num_ped)
        telefone = str(telefone)
        crm = str(observacoes)
        
        if num_ped != num_pedido_ant:
            quebrar = None
            while quebrar == None:
                #Clica no botão de Incluir Normal
                time.sleep(1)
                pgui.moveTo(450,130)
                pgui.click()
                time.sleep(1)

                #Verifica se está na tela correta
                tela_incluir = pgui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)
                
                if tela_incluir == None:
                    #Caso não esteja na tela correta, abre o smart novamente
                    funcoes_pgui.tela_inclusao(tela_incluir)
                
                #Clica em Nova Receita via outra requisição
                pgui.moveTo(228,440)
                pgui.click()
                time.sleep(0.5)
                
                #Clica no Botão OK
                pgui.moveTo(440,780)
                pgui.click()
                time.sleep(0.5)
                
                #Inclui o número da req
                pgui.moveTo(440,330)
                pgui.doubleClick()
                pgui.press('backspace')
                time.sleep(0.5)
                pgui.write(cod_item)
                pgui.press('enter')
                time.sleep(1)
                
                #Verifica se o produto é um kit
                funcoes_pgui.tratamento(cod_item)
                time.sleep(0.5)
                
                #Confirma o Item selecionado
                pgui.moveTo(1440,680)
                pgui.click()
                time.sleep(1)
                
                #Verifica se possui erro de dinamica
                dinamica = pgui.locateOnScreen('imagens/dinamica2.png', confidence=0.8)
                
                #Caso possua quebra o código
                if dinamica != None:
                    pgui.press('esc')
                    quebrar = "Dinâmica"
                    req = "Dinâmica"
                    reqs.append(req)
                    req_anterior = req
                    break
                
                #pula as telas de aviso
                pgui.moveTo(1440,720)
                pgui.click()            
                time.sleep(1)
                
                #Pula a tela de confirmação se houver
                confirmacao = pgui.locateOnScreen('imagens/confirmacao.png', confidence=0.8)
                if confirmacao == None:
                    pgui.press(['enter', 'enter'])
                else:
                    pgui.press('enter')
                
                #Armazena o número da Req
                pgui.moveTo(100,210)
                pgui.doubleClick()
                pgui.hotkey('ctrl','c')
                req = pyperclip.paste()
                req = req[:7]
                
                if req not in reqs:
                    enviar_api.consulta_api(req, num_ped)
                
                #inclui a req obtida em uma lista
                reqs.append(req)
                req_anterior = req

                #Clica em editar Req
                pgui.moveTo(700,125)
                pgui.click()
                time.sleep(0.5)
                
                #ajusta para OR e altera a data e horário de saída
                pgui.moveTo(300,210)
                pgui.click()
                time.sleep(0.5)
                pgui.write("OR")
                pgui.press('enter')
                time.sleep(0.5)
                pgui.press(['tab', 'tab'])
                
                #incluir data de produção
                pgui.write(data_producao)
                pgui.press('tab')
                pgui.write("163000")
                time.sleep(0.5)
                
                #verifica se cliente possui cadastro
                pgui.moveTo(500,270)
                pgui.click()
                time.sleep(0.5)
                pgui.moveTo(520,250)
                pgui.doubleClick()
                time.sleep(0.5)
                pgui.press('del')
                pgui.write(cpf_cli)
                pgui.press('enter')
                time.sleep(1)
                pgui.press('enter')
                time.sleep(1)
                
                #verifica se cliente possui cadastro
                button = pgui.locateOnScreen('imagens/referencia.png', confidence=0.8)

                #caso o cliente tenha cadastro
                if button != None:
                    #Sai da tela de consulta de CPF
                    pgui.press('esc')
                    time.sleep(0.5)
                    
                    #Cria o cadastro do cliente
                    funcoes_pgui.cadastro(nome_cli, cpf_cli, telefone)

                # incluir CRM
                pgui.moveTo(180,328)
                pgui.doubleClick()
                time.sleep(0.5)
                pgui.write(crm)
                pgui.press('enter')
                time.sleep(0.5)
                
                #ajusta a quantidade
                time.sleep(1)   
                pgui.moveTo(330,400)
                pgui.doubleClick()
                time.sleep(0.5)
                pgui.write(qtd)
                time.sleep(1)
                
                #calcula
                pgui.press('F9')
                time.sleep(1)
                
                #Verifica se possui mensagem de produto incompativel
                produto_incompativel = pgui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                
                pgui.press('enter')
                pgui.moveTo(x=950, y=660)
                pgui.click()
                time.sleep(1)
                
                #grava a req
                pgui.hotkey('Alt', 'G')
                time.sleep(1)
                
                #Verifica se possui mensagem de produto incompatível
                funcoes_pgui.incompativel(produto_incompativel)
                time.sleep(1)
                
                #Valida se a req está sem estoque
                sem_estoque = pgui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
                funcoes_pgui.estoque(sem_estoque)
                time.sleep(0.5)
                
                #marca como não entregar
                time.sleep(1)
                pgui.press('right')
                pgui.press('enter')
                time.sleep(1)
                pgui.moveTo(1200,730)
                pgui.click()
                time.sleep(1) 
                
                if cod_item in funcoes_pgui.kits:
                    #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial                
                    pgui.moveTo(145,135)
                    pgui.click()
                    time.sleep(0.5)
                    
                    #Clica em editar Req
                    pgui.moveTo(700,125)
                    pgui.click()
                    time.sleep(0.5)
                    
                    #ajusta a quantidade    
                    pgui.moveTo(330,400)
                    pgui.doubleClick()
                    pgui.write(qtd)
                    time.sleep(0.5)
                    
                    #calcula
                    pgui.press('F9')
                    time.sleep(1)
                    
                    #Verifica se possui mensagem de produto incompativel
                    produto_incompativel = pgui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                    time.sleep(0.5)
                    
                    #Pula as telas de confirmação
                    pgui.press('enter')
                    pgui.moveTo(x=950, y=660)
                    pgui.click()
                    time.sleep(1)
                    
                    #grava a req
                    pgui.hotkey('Alt', 'G')
                    time.sleep(0.5)
                    
                    #Verifica se possui mensagem de produto incompativel
                    funcoes_pgui.incompativel(produto_incompativel)
                    time.sleep(1)
                    
                    #Valida se a req está sem estoque
                    sem_estoque = pgui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
                    time.sleep(0.5)
                    
                    #Verifica se aparece a tela de login de sem estoque
                    funcoes_pgui.estoque(sem_estoque)
                    time.sleep(0.5)

                    #fecha a tela
                    time.sleep(1)
                    pgui.moveTo(1220,730)
                    pgui.click()
                break
        # Caso a req seja do mesmo pedido, inclui uma nova sequencial
        else:
            quebrar = None
            while quebrar == None:
                #Verifica se a primeira req teve erro de dinâmica
                if req_anterior == "Dinâmica":
                    quebrar == "Dinâmica"
                
                #Clica no botão de Incluir Normal
                pgui.moveTo(450,130)
                pgui.click()
                time.sleep(0.5)
                
                #clica em "Sequencial - Via outra requisição"
                pgui.moveTo(245,370)
                pgui.click()
                time.sleep(0.5)
                
                #Clica no Botão OK
                pgui.moveTo(440,780)
                pgui.click()
                time.sleep(0.5)
                
                #Inclui o número da req
                pgui.moveTo(440,330)
                pgui.doubleClick()
                pgui.press('backspace')
                pgui.write(cod_item)
                pgui.press('enter')
                time.sleep(1)
                
                funcoes_pgui.tratamento(cod_item)
                time.sleep(0.5)
                
                #Confirma o item selecionado
                pgui.moveTo(1440,680)
                pgui.click()
                time.sleep(1)
                
                dinamica2 = pgui.locateOnScreen("imagens/dinamica2.png", confidence=0.9)
                if dinamica2 != None:
                    pgui.press("enter")
                    quebrar = "Dinâmica"
                    req = "SEQ-Dinamica"
                    break
                
                #Pula a tela de aviso
                pgui.moveTo(1440,720)
                pgui.click()            
                time.sleep(1)
                
                #Clica em editar Req
                pgui.moveTo(700,125)
                pgui.click()
                time.sleep(0.5)

                # incluir CRM
                pgui.moveTo(180,328)
                pgui.doubleClick()
                time.sleep(0.5)
                pgui.write(crm)
                pgui.press('enter')
                time.sleep(0.5)
                
                #ajusta a quantidade    
                pgui.moveTo(330,400)
                pgui.doubleClick()
                pgui.write(qtd)
                time.sleep(0.5)
                
                #calcula
                pgui.press('F9')
                time.sleep(1)
                
                produto_incompativel = pgui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                
                #Pula a tela de produto incompatível
                pgui.press('enter')
                pgui.moveTo(x=950, y=660)
                pgui.click()
                time.sleep(1)
                
                #grava a req
                pgui.hotkey('Alt', 'G')
                time.sleep(0.5)
                
                funcoes_pgui.incompativel(produto_incompativel)
                time.sleep(0.5)
                
                #Valida se a req está sem estoque
                sem_estoque = pgui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
                time.sleep(0.5)

                funcoes_pgui.estoque(sem_estoque)
                time.sleep(0.5)
                
                #fecha a tela
                pgui.moveTo(1220,730)
                pgui.click()
                time.sleep(0.5)
                
                if cod_item in funcoes_pgui.kits:
                    #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial
                    pgui.moveTo(145,135)
                    pgui.click()
                    time.sleep(0.5)
                    
                    #Clica em editar Req
                    pgui.moveTo(700,125)
                    pgui.click()
                    time.sleep(0.5)
                    
                    #ajusta a quantidade    
                    pgui.moveTo(330,400)
                    pgui.doubleClick()
                    pgui.write(qtd)
                    time.sleep(0.5)
                    
                    #calcula
                    pgui.press('F9')
                    time.sleep(1)
                    
                    #Verifica se possui mensagem de produto incompativel
                    produto_incompativel = pgui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                    time.sleep(0.5)
                    
                    #Pula a tela de produto incompatível e 
                    pgui.press('enter')
                    pgui.moveTo(x=950, y=660)
                    pgui.click()
                    time.sleep(1)
                    
                    #grava a req
                    pgui.hotkey('Alt', 'G')
                    time.sleep(0.5)
                    
                    #Verifica se possui mensagem de produto incompativel
                    funcoes_pgui.incompativel(produto_incompativel)
                    time.sleep(0.5)
                    
                    #Valida se a req está sem estoque
                    sem_estoque = pgui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
                    time.sleep(0.5)
                    
                    #Verifica se possui mensagem de sem estoque
                    funcoes_pgui.estoque(sem_estoque)
                    time.sleep(0.5)
                    
                    #fecha a tela
                    time.sleep(1)
                    pgui.moveTo(1220,730)
                    pgui.click()
                    time.sleep(0.5)
                break
        #Armazena o número do pedido como o anterior para verificação
        num_pedido_ant = num_ped
        
        #Armazena a hora para criação de log
        hora = datetime.now()
        
        #Criação de log
        log = "\nHora: {}, Pedido: {}, Cliente: {}, CPF Cliente: {} ,REQ: {}, Data Smart: {}".format(hora, num_ped, nome_cli, cpf_cli, req, data_producao)
        with open(f'log/log{data_producao}.txt', 'a') as arquivo:
            arquivo.write(str(log))
    
    janela_final = tk.Tk()
    janela_final.title("Finalizado")
    janela_final.geometry("1330x200")
    tk.Label(janela_final, text="Pedidos Incluídos", font=("Helvetica", 90, "bold")).place(x=10, y=20)
    janela_final.wm_attributes('-topmost', True)
    janela_final.mainloop()

#Rótulo segundo passo
tk.Label(root, text="2º Unir os pedidos").place(x=50,y=100)

#Botão para unir os pedidos
button_unir = tk.Button(root, text="Unir Pedidos", command=combine_excel_files)
button_unir.place(x=250, y=130)

#Rótulos de quantidades de pedidos com base na analise dos pedidos
total_pedidos_label = tk.Label(root, text="")
total_pedidos_label.place(x=50, y=150)

#Rótulos de quantidades de reqs com base na analise dos pedidos
total_requisicoes_label = tk.Label(root, text="")
total_requisicoes_label.place(x=50, y=170)

#Rótulo 3º passo
tk.Label(root, text="3º Selecione a data para produção").place(x=50,y=220)

#Obtém a data do dia
data = date.today()

#Armazena a data e o ano
ano = data.year
mes = data.month

#Adiciona o calendário
cal = Calendar(root, selectmode = 'day', year = ano, month = mes)
cal.place(x=50, y=250)

#Função para obter a data selecionada
def grad_date():
    global data_producao
    #Armazena a data selecionada
    data_selecionada = cal.get_date()
    
    #Separa a data para melhor organização
    mes, dia, ano = data_selecionada.split("/")
    
    if len(mes) <2:
        mes = "0" + mes
    #Ajusta a variável para atender a data de produção
    data_producao = dia + mes + "20" + ano
    
    #Ajusta a data para mostrar ao usuário a data selecionada
    data_view = dia + "/" + mes + "/20" + ano
    
    #Condição para incluir mais um caractere no dia caso não possua
    if len(data_producao) < 8:
        data_producao = "0" + data_producao
    
    #Rótulo para visualização da data
    date.config(text = "Data: " + data_view)

# Adiciona o botão para definir a data
tk.Button(root, text = "Definir data", command = grad_date).place(x=250, y=470)

#Rótulo para exibir a data selecionada
date = tk.Label(root, text = "")
date.place(x=50, y=480)

#Rótulo 4º passo
tk.Label(root, text="4º Iniciar inclusão").place(x=50,y=520)

#Botão para unir os pedidos
button_unir = tk.Button(root, text="Iniciar", command=incluir_smart)
button_unir.place(x=250, y=550)

root.mainloop()
