import pandas as pd
import xlrd as xlrd
import pyautogui as pgui
import numpy as np
import tkinter as tk
import os
import time
import requests
import pyperclip
import json
import cv2
from tkinter import filedialog
from tkcalendar import Calendar
from unidecode import unidecode

#Criação da interface gráfica no tkinter
root = tk.Tk()

root.title("Incluir Pedidos")

root.geometry("400x650")

root.resizable(0, 0)

canvas = tk.Canvas(root, width=200, height=200)

canvas.pack()

#Função para armazenar o local da pasta selecionado
def proc_pasta():
    global pasta_path
    #Armazena o caminho da pasta selecionada
    pasta_path = filedialog.askdirectory()

#Rótulo da tela
tk.Label(root, text="1º Selecione a pasta com os pedidos").place(x=50,y=30)

#Botão para procurar a pasta
button = tk.Button(root, text="Procurar", command=proc_pasta)
button.place(x=250, y=60)

#Função para juntar os arquivos da pasta em uma única planilha
def combine_excel_files():
    global separados
    # Lista para armazenar os dataframes
    df_list = []
    folder_path = pasta_path
    output_file = "pedidos.xlsx"
    # Percorre todos os arquivos na pasta
    for file in os.listdir(folder_path):
        
        # Verifica se o arquivo é um arquivo Excel
        if file.endswith('.xlsx') or file.endswith('.xls'):
            
            # Lê o arquivo Excel para um dataframe
            df = pd.read_excel(os.path.join(folder_path, file))
            
            # Adiciona o dataframe à lista
            df_list.append(df)

    # Combina todos os dataframes
    combined_df = pd.concat(df_list)

    # Escreve o dataframe combinado para um novo arquivo Excel
    combined_df.to_excel(folder_path + "/" + output_file, index=False)
    
    #Consulta os pedidos Semi-Acabados no banco de dados
    response = requests.get("https://api.fmiligrama.com/produtos?tipoInterno=Semi-Acabado").json()

    #Armazena os itens Semi-Acabados em um dataframe
    semiacabados = pd.json_normalize(response['data'])

    #Cria um Data Frame com a planilha de pedidos
    pedidos = pd.read_excel(folder_path + "/" + output_file, sheet_name='Sheet1')

    #Agrupa os itens duplicados
    pedidos = pedidos.groupby(["Número do pedido","Fone", "CPF/CNPJ", "Nome do contato", "Código (SKU)"]).sum("Quantidade")
    pedidos = pedidos.reset_index()

    #Converte a coluna codigo da base de dados para inteiro
    semiacabados["codigo"]=semiacabados["codigo"].astype(int)

    #Mescla a planilha de pedidos com a base de dados para filtrar os pedidos semi-acabados
    separados = pd.merge(pedidos, semiacabados, left_on="Código (SKU)", right_on="codigo", how="left")

    #Acrescenta a descrição de manipulado nos itens que estão com tipo interno em branco
    separados["tipoInterno"] = separados["tipoInterno"].replace(np.nan,"Manipulado")

    #Ajusta o CPF do cliente para que atenda os requisitos do smart
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace(".","", regex=False)
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace("-","", regex=False)
    separados["CPF/CNPJ"] = separados["CPF/CNPJ"].str.replace("/","", regex=False)

    #Ajusta o telefone do cliente para que atenda os requisitos do smart
    separados["Fone"] = separados["Fone"].str.replace(")","", regex=False)
    separados["Fone"] = separados["Fone"].str.replace("(","", regex=False)
    separados["Fone"] = separados["Fone"].str.replace("-","", regex=False)
    separados.loc[separados["Fone"].str.contains("x"), "Fone"] = 99999999999
    
    #Ajusta os códigos para que fiquem sem o 99 na frente
    sem_ajuste = ["996760", "995066", "994409"]
    separados["Código (SKU)"] = separados["Código (SKU)"].astype(str)
    for i in range(len(separados["Código (SKU)"])):
        if separados["Código (SKU)"][i] in sem_ajuste:
            print(i)
        else:
            if separados["Código (SKU)"][i].startswith('99'):
                separados["Código (SKU)"][i] = separados["Código (SKU)"][i].replace('99','',1)
    
    separados["Nome do contato"] = separados["Nome do contato"].apply(unidecode)

    #Armazena os pedidos manipulados
    manipulados = separados[separados["tipoInterno"].isin(["Manipulado"])]
    
    #Soma a quantidade de reqs e pedidos
    total_pedidos = manipulados["Número do pedido"].drop_duplicates()
    total_pedidos = total_pedidos.count()
    total_requisicoes = manipulados["Quantidade"].sum()
    
    #Cria os rótulos para serem mostrados ao usuário
    total_pedidos_label.config(text=f"Total Pedidos: {total_pedidos}")
    total_requisicoes_label.config(text=f"Total REQs: {total_requisicoes}")

def incluir_smart():
    #Separa todos os itens que precisam ser inclusos no smart
    smart = separados["tipoInterno"] == "Manipulado"
    smart = separados.loc[smart]
    
    #Variável para armazenar os números de pedidos
    num_pedido_ant = None
    num_pedido_ant = str(num_pedido_ant)
    
    #Lista para verificação de reqs no sistema
    requisicoes = []
    
    #Função para incluir os pedidos no smart
    for i in range(len(smart)):        
        #Define uma variável para cada valor a ser usado no código
        cod_item = smart["Código (SKU)"].iloc[i]
        nome_cli = smart["Nome do contato"].iloc[i]
        cpf_cli = smart["CPF/CNPJ"].iloc[i]
        num_ped = smart["Número do pedido"].iloc[i]
        qtd = smart["Quantidade"].iloc[i]
        telefone = smart["Fone"].iloc[i]
        cod_item = str(cod_item)
        qtd = qtd = str(qtd)
        num_ped = str(num_ped)
        telefone = str(telefone)
        
        if num_ped != num_pedido_ant:
            #Clica no botão de Incluir Normal
            time.sleep(1)
            pgui.moveTo(450,130)
            pgui.click()
            
            #Clica em Nova Receita via outra requisição
            time.sleep(1)
            pgui.moveTo(228,440)
            pgui.click()
            
            #Clica no Botão OK
            pgui.moveTo(440,780)
            pgui.click()
            
            #Inclui o número da req
            pgui.moveTo(440,330)
            pgui.doubleClick()
            pgui.press('backspace')
            pgui.write(cod_item)
            pgui.press('enter')
            time.sleep(1)
            
            #Verifica se o item é tratamento para unhas
            if cod_item in ["430035","396864", "702310"]:
                pgui.moveTo(570,430)
                pgui.doubleClick()
                pgui.moveTo(570,470)
                pgui.doubleClick()
            else:
                pgui.moveTo(570,430)
                pgui.doubleClick()
            
            #pula as telas de aviso
            pgui.moveTo(1440,680)
            pgui.click()
            time.sleep(1)
            pgui.moveTo(1440,720)
            pgui.click()            
            time.sleep(2)
            
            #Pula a tela de confirmação se houver
            confirmacao = pgui.locateOnScreen('confirmacao.png', confidence=0.8)
            if confirmacao == None:
                pgui.press(['enter', 'enter'])
            else:
                pgui.press('enter')
            
            #Armazena o número da Req
            pgui.moveTo(100,210)
            pgui.doubleClick()
            pgui.hotkey('ctrl','c')
            req = pyperclip.paste()
            
            if req not in requisicoes:
                #Api tiny
                url = "https://api.fmiligrama.com/vendas/smart"
                num_ped = str(num_ped)
                
                #Dados para api
                payload = json.dumps({
                    "numero_tiny": num_ped,
                    "req_smart": req
                })
                #Cabeçalho para envio da api
                headers = {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                }
                #Comando de envio para api
                response = requests.request("POST", url, headers=headers, data=payload)
            
            requisicoes.append(req)
            #Clica em editar Req
            pgui.moveTo(700,125)
            pgui.click()

            #ajusta para OR e altera a data e horário de saída
            pgui.moveTo(300,210)
            pgui.click()
            pgui.write("OR")
            pgui.press('enter')
            pgui.press(['tab', 'tab'])

            #incluir data de produção
            pgui.write(data_producao)
            pgui.press('tab')
            pgui.write("110000")
            
            #verifica se cliente possui cadastro
            pgui.moveTo(500,270)
            pgui.click()
            pgui.moveTo(520,250)
            pgui.doubleClick()
            pgui.press('del')
            pgui.write(cpf_cli)
            pgui.press('enter')
            time.sleep(1)
            pgui.press('enter')
            time.sleep(1)

            #verifica se cliente possui cadastro
            button = pgui.locateOnScreen('referencia.png', confidence=0.8)
            #caso o cliente tenha cadastro
            if button != None:
                pgui.press('esc')
                
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
            
            #ajusta a quantidade
            time.sleep(1)   
            pgui.moveTo(330,400)
            pgui.doubleClick()
            pgui.write(qtd)
            time.sleep(1)
            
            #calcula
            pgui.press('F9')
            time.sleep(1)
            
            #Verifica se possui mensagem de produto incompativel
            produto_incompativel = pgui.locateOnScreen('prod_incomp.png', confidence= 0.8)
            
            pgui.press('enter')
            pgui.moveTo(x=950, y=660)
            pgui.click()
            time.sleep(1)
            
            #grava a req
            pgui.hotkey('Alt', 'G')
            time.sleep(1)
            
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
            time.sleep(1)
            
            #Valida se a req está sem estoque
            sem_estoque = pgui.locateOnScreen('sem-estoque.png', confidence= 0.8)
            if sem_estoque != None:
                pgui.moveTo(940,550)
                pgui.click()
                pgui.write("ALLAN.J")
                pgui.press('tab')
                pgui.press('tab')
                pgui.write("01825675")
                pgui.press('enter')
            
            #marca como não entregar
            time.sleep(1)
            pgui.press('right')
            pgui.press('enter')
            time.sleep(1)
            pgui.moveTo(1200,730)
            pgui.click()
            time.sleep(1) 
            
            if cod_item in ["430035","396864", "702310"]:
                #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial                
                pgui.moveTo(145,135)
                pgui.click()
                
                #Clica em editar Req
                pgui.moveTo(700,125)
                pgui.click()
                
                #ajusta a quantidade    
                pgui.moveTo(330,400)
                pgui.doubleClick()
                pgui.write(qtd)
                
                #calcula
                pgui.press('F9')
                time.sleep(1)
                
                #Verifica se possui mensagem de produto incompativel
                produto_incompativel = pgui.locateOnScreen('prod_incomp.png', confidence= 0.8)
                
                #Pula as telas de confirmação
                pgui.press('enter')
                pgui.moveTo(x=950, y=660)
                pgui.click()
                time.sleep(1)
                
                #grava a req
                pgui.hotkey('Alt', 'G')
                if produto_incompativel == None:                
                    time.sleep(1)
                    pgui.press('enter')
                else:
                    time.sleep(1)
                    pgui.press('enter')
                    time.sleep(1)
                    pgui.press('enter')
                time.sleep(1)

                #Valida se a req está sem estoque
                sem_estoque = pgui.locateOnScreen('sem-estoque.png', confidence= 0.8)
                if sem_estoque != None:
                    pgui.moveTo(940,550)
                    pgui.click()
                    pgui.write("MILIBOT")
                    pgui.press('tab')
                    pgui.press('tab')
                    pgui.write("77777777")
                    pgui.press('enter')
                    
                #fecha a tela
                time.sleep(1)
                pgui.moveTo(1220,730)
                pgui.click()
        # Caso a req seja do mesmo pedido, inclui uma nova sequencial
        else:
            #Clica no botão de Incluir Normal
            pgui.moveTo(450,130)
            pgui.click()
            
            #clica em "Sequencial - Via outra requisição"
            pgui.moveTo(245,370)
            pgui.click()
            
            #Clica no Botão OK
            pgui.moveTo(440,780)
            pgui.click()
            
            #Inclui o número da req
            pgui.moveTo(440,330)
            pgui.doubleClick()
            pgui.press('backspace')
            pgui.write(cod_item)
            pgui.press('enter')
            time.sleep(1)
            
            #Verifica se o item é tratamento para unhas
            if cod_item in ["430035","396864", "702310"]:
                pgui.moveTo(570,430)
                pgui.doubleClick()
                pgui.moveTo(570,470)
                pgui.doubleClick()
            else:
                pgui.moveTo(570,430)
                pgui.doubleClick()
            
            #pula as telas de aviso
            pgui.moveTo(1440,680)
            pgui.click()
            time.sleep(1)
            pgui.moveTo(1440,720)
            pgui.click()            
            time.sleep(1)
            
            #Clica em editar Req
            pgui.moveTo(700,125)
            pgui.click()
            
            #ajusta a quantidade    
            pgui.moveTo(330,400)
            pgui.doubleClick()
            pgui.write(qtd)
            
            #calcula
            pgui.press('F9')
            time.sleep(1)
            
            produto_incompativel = pgui.locateOnScreen('prod_incomp.png', confidence= 0.8)
            
            #Pula a tela de produto incompatível
            pgui.press('enter')
            pgui.moveTo(x=950, y=660)
            pgui.click()
            time.sleep(1)
            
            #grava a req
            pgui.hotkey('Alt', 'G')
            
            if produto_incompativel == None:                
                time.sleep(1)
                pgui.press('enter')
            else:
                time.sleep(1)
                pgui.press('enter')
                time.sleep(1)
                pgui.press('enter')
            time.sleep(1)
            
            #Valida se a req está sem estoque
            sem_estoque = pgui.locateOnScreen('sem-estoque.png', confidence= 0.8)
            if sem_estoque != None:
                pgui.moveTo(940,550)
                pgui.click()
                pgui.write("MILIBOT")
                pgui.press('tab')
                pgui.press('tab')
                pgui.write("77777777")
                pgui.press('enter')
                
            #fecha a tela
            pgui.moveTo(1220,730)
            pgui.click()
            
            if cod_item in ["430035","396864", "702310"]:
                #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial
                pgui.moveTo(145,135)
                pgui.click()
                
                #Clica em editar Req
                pgui.moveTo(700,125)
                pgui.click()
                
                #ajusta a quantidade    
                pgui.moveTo(330,400)
                pgui.doubleClick()
                pgui.write(qtd)
                
                #calcula
                pgui.press('F9')
                time.sleep(1)
                
                #Verifica se possui mensagem de produto incompativel
                produto_incompativel = pgui.locateOnScreen('prod_incomp.png', confidence= 0.8)
                
                #Pula a tela de produto incompatível e 
                pgui.press('enter')
                pgui.moveTo(x=950, y=660)
                pgui.click()
                time.sleep(1)
                
                #grava a req
                pgui.hotkey('Alt', 'G')
                if produto_incompativel == None:                
                    time.sleep(1)
                    pgui.press('right')
                    pgui.press('enter')
                else:
                    time.sleep(1)
                    pgui.press('enter')
                    time.sleep(1)
                    pgui.press('enter')
                time.sleep(1)

                #Valida se a req está sem estoque
                sem_estoque = pgui.locateOnScreen('sem-estoque.png', confidence= 0.8)
                if sem_estoque != None:
                    pgui.moveTo(940,550)
                    pgui.click()
                    pgui.write("MILIBOT")
                    pgui.press('tab')
                    pgui.press('tab')
                    pgui.write("77777777")
                    pgui.press('enter')
                    
                #fecha a tela
                time.sleep(1)
                pgui.moveTo(1220,730)
                pgui.click()
        #Armazena o número do pedido como o anterior para verificação
        num_pedido_ant = num_ped

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

#Adiciona o calendário
cal = Calendar(root, selectmode = 'day', year = 2024, month = 1)
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