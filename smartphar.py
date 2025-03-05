import time
import pandas as pd
import pyautogui as pgui
from order import Order
from locations import *


def open_smartphar():
    # Windows Run
    pgui.hotkey("win", "r")
    pgui.write(r"\\10.1.1.5\j\smartpharmacy\SmartPhar.exe")
    pgui.press("enter")
    time.sleep(2)
    pgui.press("left")
    pgui.press("enter")
    time.sleep(5)


def login_smartphar():
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


def open_receitas_screen():
    receita_icon = pgui.locateOnScreen("imagens/receita_smart.png", confidence=0.5)
    if receita_icon != None:
        pgui.moveTo(receita_icon_location['x'], receita_icon_location['y'])
        pgui.click()


def filter_manipulados(folder_path):
    # read filtered_orders.xslx file
    filtered_orders = pd.read_excel(folder_path + "/" + 'filtered_orders.xlsx', sheet_name='Sheet1')

    # Filter only the orders that need to be included in smartphar
    smart_filtered_orders = filtered_orders["tipoInterno"] == "manipulado"
    smart_filtered_orders = filtered_orders.loc[smart_filtered_orders]

    return smart_filtered_orders


def click_incluir_normal():
    pgui.moveTo(incluir_icon_location['x'], incluir_icon_location['y'])
    pgui.click()
    time.sleep(1)


def click_nova_receita_outra_requisicao():
    pgui.moveTo(nova_receita_via_outra_requisicao_location['x'], nova_receita_via_outra_requisicao_location['y'])
    pgui.click()
    time.sleep(1)
    pgui.press("enter")


def pesquisar_requisicao_inclusao_via_outra_receita():
    # TODO: Implement this function
    pgui.moveTo(440,330)
    pgui.doubleClick()
    pgui.press('backspace')
    time.sleep(0.5)
    pgui.write(cod_item)
    pgui.press('enter')
    time.sleep(1)


def insert_orders_smartphar(smart_filtered_orders):
    # Variables to receive order number and reqs
    previous_order_number = None
    previous_order_number = str(previous_order_number)
    reqs = []
    
    for i in range(len(smart_filtered_orders)):        
        #Define uma variável para cada valor a ser usado no código
        sku = smart_filtered_orders["Código (SKU)"].iloc[i]
        name = smart_filtered_orders["Nome do contato"].iloc[i]
        cpf = smart_filtered_orders["CPF/CNPJ"].iloc[i]
        order_number = smart_filtered_orders["Número do pedido"].iloc[i]
        qtd = smart_filtered_orders["Quantidade"].iloc[i]
        phone = smart_filtered_orders["Fone"].iloc[i]
        
        order = Order(sku, name, cpf, order_number, qtd, phone)
        print(order)
        
        # if order[order_number] != previous_order_number:
        while True:
            # Clica no botão de Incluir Normal
            time.sleep(1)
            click_incluir_normal()                

            # Verifica se está na tela correta
            tela_incluir = pgui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)
            
            # Caso não esteja na tela correta, abre o smart novamente
            if tela_incluir == None:
                open_smartphar()
                login_smartphar()
                open_receitas_screen()
            else:
                break

        # Clica em Nova Receita via outra requisição
        click_nova_receita_outra_requisicao()

        pesquisar_requisicao_inclusao_via_outra_receita()

            
        break
                
                

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
                time.sleep(0.5)
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
        
    #     #Armazena a hora para criação de log
    #     hora = datetime.now()
        
    #     #Criação de log
    #     log = "\nHora: {}, Pedido: {}, Cliente: {}, CPF Cliente: {} ,REQ: {}, Data Smart: {}".format(hora, num_ped, nome_cli, cpf_cli, req, data_producao)
    #     with open(f'log/log{data_producao}.txt', 'a') as arquivo:
    #         arquivo.write(str(log))
    
    # janela_final = tk.Tk()
    # janela_final.title("Finalizado")
    # janela_final.geometry("1330x200")
    # tk.Label(janela_final, text="Pedidos Incluídos", font=("Helvetica", 90, "bold")).place(x=10, y=20)
    # janela_final.wm_attributes('-topmost', True)
    # janela_final.mainloop()
