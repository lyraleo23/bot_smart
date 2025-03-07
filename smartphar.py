import time
import pandas as pd
import pyautogui
from order import Order
from locations import *

pyautogui.PAUSE = 0.2


def open_smartphar():
    # Windows Run
    pyautogui.hotkey("win", "r")
    pyautogui.write(r"\\10.1.1.5\j\smartpharmacy\SmartPhar.exe")
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("left")
    pyautogui.press("enter")
    time.sleep(5)


def login_smartphar(production_branch):
    # Insert login
    pyautogui.write("MILIBOT")
    pyautogui.press('TAB')
    pyautogui.press('TAB')

    # Insert password
    pyautogui.write('77777777')
    pyautogui.press('TAB')

    # Select branch
    if production_branch == '100':
        pyautogui.write
        for _ in range(5):
            pyautogui.press('down')
        pyautogui.press('enter')
    elif production_branch == '600':
        pyautogui.write
        pyautogui.press('down')
        pyautogui.press('enter')
    
    time.sleep(2)
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')


def open_receitas_screen():
    # receita_icon = pyautogui.locateOnScreen("imagens/receita_smart.png", confidence=0.5)
    # if receita_icon != None:
    #     pyautogui.moveTo(receita_icon_location['x'], receita_icon_location['y'])
    #     pyautogui.click()
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('r')
    time.sleep(0.1)
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('r')


def filter_manipulados(folder_path):
    # read filtered_orders.xslx file
    filtered_orders = pd.read_excel(folder_path + "/" + 'filtered_orders.xlsx', sheet_name='Sheet1')

    # Filter only the orders that need to be included in smartphar
    smart_filtered_orders = filtered_orders["tipoInterno"] == "manipulado"
    smart_filtered_orders = filtered_orders.loc[smart_filtered_orders]

    return smart_filtered_orders


def click_incluir_normal():
    pyautogui.hotkey('alt', 'i')


def click_nova_receita_via_outra_requisicao():
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("enter")


def click_sequencial_outra_requisicao():
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("enter")


def verificar_kit(sku):
    kits_list = ['430035', '396864', '702310']
    if sku in kits_list:
        return True
    else:
        return False


def pesquisar_requisicao_inclusao_via_outra_receita(sku):
    pyautogui.moveTo(fied_receita_via_outra_requisicao_numero_req_location['x'], fied_receita_via_outra_requisicao_numero_req_location['y'])
    pyautogui.doubleClick()
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.write(sku)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)


def verify_error_dinamica():
    try:
        error_dinamica = pyautogui.locateOnScreen('imagens/erro_dinamica.PNG', confidence=0.8)
    except:
        error_dinamica = None
    
    return error_dinamica


def click_alterar():
    pyautogui.hotkey('alt', 'a')
    pyautogui.press('tab')


def transformar_or():
    pyautogui.write('OR')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')


def altera_data_hora_entrega(production_date):
    # Altera data
    pyautogui.write(production_date)
    pyautogui.press('tab')
    
    # Hora de entrega fix em 16:30
    pyautogui.write('163000')
    time.sleep(0.5)


def search_customer(cpf):
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(cpf)
    pyautogui.press('enter')
    time.sleep(1)

    #verifica se cliente possui cadastro
    try:
        customer_verification = pyautogui.locateOnScreen('imagens/referencia.png', confidence=0.8)
    except:
        customer_verification = None

    if customer_verification == None:
        pyautogui.press('enter')
    else:
        pyautogui.press('esc')
        time.sleep(0.2)

    return customer_verification

        
def cadastrar_cliente(name, cpf, phone):
    pyautogui.press('insert')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write(name)
    for _ in range(4):
        pyautogui.press('tab')
    pyautogui.write(cpf)
    for _ in range(6):
        pyautogui.press('tab')
    pyautogui.write('M')
    for _ in range(9):
        pyautogui.press('tab')
    pyautogui.write(phone)
    
    pyautogui.hotkey('alt', 'g')


def incluir_crm(crm):
    for _ in range(9):
        pyautogui.press('tab')
    pyautogui.write(crm)
    pyautogui.press('enter')


def ajustar_quantidade(qtd):
    for _ in range(12):
        pyautogui.press('tab')
    pyautogui.write(qtd)
    pyautogui.press('F9')
    time.sleep(0.5)
    for _ in range(3):
        pyautogui.press('enter')


def save_req():
    pyautogui.hotkey('alt', 'g')
    pyautogui.press('enter')
    pyautogui.press('right')
    pyautogui.press('enter')


def ajustar_quantidade_manual(qtd):
    for _ in range(3):
        pyautogui.press('tab')
    pyautogui.write(qtd)
    pyautogui.press('F9')
    time.sleep(0.5)
    for _ in range(3):
        pyautogui.press('enter')


def insert_orders_smartphar(smart_filtered_orders, sector_var, production_branch, production_date):
    # Variables to receive order number and reqs
    previous_order_number = None
    error = None
    reqs = []
    orders_errors = []
    
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
        print(f'ERROR entrada: {error}')

        if error == 'dinamica':
            if order.order_number == previous_order_number:
                continue
            else:
                error = None
        
        print(f'ERROR saida: {error}') 

        if order.order_number != previous_order_number:
            while True:
                try:
                    # Clica no botão de Incluir Normal
                    time.sleep(1)
                    click_incluir_normal()
                    time.sleep(2)

                    # Verifica se está na tela correta
                    tela_incluir = pyautogui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)
                except Exception as e:
                    print('Erro')
                    tela_incluir = None

                # Caso não esteja na tela correta, abre o smart novamente
                if tela_incluir == None:
                    open_smartphar()
                    login_smartphar(production_branch)
                    open_receitas_screen()
                else:
                    break

            # Clica em Nova Receita via outra requisição
            click_nova_receita_via_outra_requisicao()
            time.sleep(1)
            
        elif order.order_number == previous_order_number:
            # Verifica se está na tela correta
            tela_incluir = pyautogui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)

            if tela_incluir == None:
                open_smartphar()
                login_smartphar(production_branch)
                open_receitas_screen()
                continue

            click_sequencial_outra_requisicao()
            time.sleep(1)

        # Clica em Nova Receita via outra requisição
        print(f'incluindo uma sequencial pedido: {order.order_number}: {order.sku}')
        pesquisar_requisicao_inclusao_via_outra_receita(order.sku)
        time.sleep(5)
        previous_order_number = order.order_number
        
        #Verifica se possui erro de dinamica
        error_dinamica = verify_error_dinamica()

        if error_dinamica != None:
            print('ERROR: dinamica!')
            error = 'dinamica'
            e = {
                'order_number': order.order_number,
                'error': error
            }
            orders_errors.append(e)

            # Fecha a tela de erro
            pyautogui.press('esc')
            continue
        
        # Confirma item
        pyautogui.press('enter')
        pyautogui.press('enter')
        time.sleep(2)

        # Clica em alterar req
        click_alterar()

        if sector_var == 'Manual':
            crm = smart_filtered_orders["Observações"].iloc[i]
            incluir_crm(str(crm))
            if qtd > 1:
                ajustar_quantidade_manual(order.qtd)
        else:
            if qtd > 1:
                ajustar_quantidade(order.qtd)
        
        save_req()
        print('Sequencial incluída com sucesso!')
        time.sleep(2)

        next_order_number = smart_filtered_orders["Número do pedido"].iloc[i+1]
        if order.order_number != next_order_number:
            click_alterar()
            transformar_or()
            altera_data_hora_entrega(production_date)
            pyautogui.press('tab')
            pyautogui.press('tab')
            customer_verification = search_customer(order.cpf)

            if customer_verification != None:
                cadastrar_cliente(order.name, order.cpf, order.phone)

            

                
                

            
                
                

        #         #Inclui o número da req
        #         pyautogui.moveTo(440,330)
        #         pyautogui.doubleClick()
        #         pyautogui.press('backspace')
        #         time.sleep(0.5)
        #         pyautogui.write(cod_item)
        #         pyautogui.press('enter')
        #         time.sleep(1)
                
        #         #Verifica se o produto é um kit
        #         funcoes_pyautogui.tratamento(cod_item)
        #         time.sleep(0.5)
                
        #         #Confirma o Item selecionado
        #         pyautogui.moveTo(1440,680)
        #         pyautogui.click()
        #         time.sleep(1)
                
        #         #Verifica se possui erro de dinamica
        #         dinamica = pyautogui.locateOnScreen('imagens/dinamica2.png', confidence=0.8)
                
        #         #Caso possua quebra o código
        #         if dinamica != None:
        #             pyautogui.press('esc')
        #             quebrar = "Dinâmica"
        #             req = "Dinâmica"
        #             reqs.append(req)
        #             req_anterior = req
        #             break
                
        #         #pula as telas de aviso
        #         pyautogui.moveTo(1440,720)
        #         pyautogui.click()            
        #         time.sleep(1)
                
        #         #Pula a tela de confirmação se houver
        #         confirmacao = pyautogui.locateOnScreen('imagens/confirmacao.png', confidence=0.8)
        #         if confirmacao == None:
        #             pyautogui.press(['enter', 'enter'])
        #         else:
        #             pyautogui.press('enter')
                
        #         #Armazena o número da Req
        #         pyautogui.moveTo(100,210)
        #         pyautogui.doubleClick()
        #         pyautogui.hotkey('ctrl','c')
        #         req = pyperclip.paste()
        #         req = req[:7]
                
        #         if req not in reqs:
        #             enviar_api.consulta_api(req, num_ped)
                
        #         #inclui a req obtida em uma lista
        #         reqs.append(req)
        #         req_anterior = req

        #         #Clica em editar Req
        #         pyautogui.moveTo(700,125)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         #ajusta para OR e altera a data e horário de saída
        #         pyautogui.moveTo(300,210)
        #         pyautogui.click()
        #         time.sleep(0.5)
        #         pyautogui.write("OR")
        #         pyautogui.press('enter')
        #         time.sleep(0.5)
        #         pyautogui.press(['tab', 'tab'])
                
        #         #incluir data de produção
        #         pyautogui.write(data_producao)
        #         pyautogui.press('tab')
        #         pyautogui.write("163000")
        #         time.sleep(0.5)
                
        #         #verifica se cliente possui cadastro
        #         pyautogui.moveTo(500,270)
        #         pyautogui.click()
        #         time.sleep(0.5)
        #         pyautogui.moveTo(520,250)
        #         pyautogui.doubleClick()
        #         time.sleep(0.5)
        #         pyautogui.press('del')
        #         pyautogui.write(cpf_cli)
        #         pyautogui.press('enter')
        #         time.sleep(1)
        #         pyautogui.press('enter')
        #         time.sleep(1)
                
        #         #verifica se cliente possui cadastro
        #         button = pyautogui.locateOnScreen('imagens/referencia.png', confidence=0.8)
                
        #         #caso o cliente tenha cadastro
        #         if button != None:
        #             #Sai da tela de consulta de CPF
        #             pyautogui.press('esc')
        #             time.sleep(0.5)
                    
        #             #Cria o cadastro do cliente
        #             funcoes_pyautogui.cadastro(nome_cli, cpf_cli, telefone)
                
        #         #ajusta a quantidade
        #         time.sleep(1)   
        #         pyautogui.moveTo(330,400)
        #         pyautogui.doubleClick()
        #         time.sleep(0.5)
        #         pyautogui.write(qtd)
        #         time.sleep(1)
                
        #         #calcula
        #         pyautogui.press('F9')
        #         time.sleep(1)
                
        #         #Verifica se possui mensagem de produto incompativel
        #         produto_incompativel = pyautogui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                
        #         pyautogui.press('enter')
        #         pyautogui.moveTo(x=950, y=660)
        #         pyautogui.click()
        #         time.sleep(1)
                
        #         #grava a req
        #         pyautogui.hotkey('Alt', 'G')
        #         time.sleep(1)
                
        #         #Verifica se possui mensagem de produto incompatível
        #         funcoes_pyautogui.incompativel(produto_incompativel)
        #         time.sleep(1)
                
        #         #Valida se a req está sem estoque
        #         sem_estoque = pyautogui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
        #         funcoes_pyautogui.estoque(sem_estoque)
        #         time.sleep(0.5)
                
        #         #marca como não entregar
        #         time.sleep(1)
        #         pyautogui.press('right')
        #         pyautogui.press('enter')
        #         time.sleep(1)
        #         pyautogui.moveTo(1200,730)
        #         pyautogui.click()
        #         time.sleep(1) 
                
        #         if cod_item in funcoes_pyautogui.kits:
        #             #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial                
        #             pyautogui.moveTo(145,135)
        #             pyautogui.click()
        #             time.sleep(0.5)
                    
        #             #Clica em editar Req
        #             pyautogui.moveTo(700,125)
        #             pyautogui.click()
        #             time.sleep(0.5)
                    
        #             #ajusta a quantidade    
        #             pyautogui.moveTo(330,400)
        #             pyautogui.doubleClick()
        #             pyautogui.write(qtd)
        #             time.sleep(0.5)
                    
        #             #calcula
        #             pyautogui.press('F9')
        #             time.sleep(1)
                    
        #             #Verifica se possui mensagem de produto incompativel
        #             produto_incompativel = pyautogui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
        #             time.sleep(0.5)
                    
        #             #Pula as telas de confirmação
        #             pyautogui.press('enter')
        #             pyautogui.moveTo(x=950, y=660)
        #             pyautogui.click()
        #             time.sleep(1)
                    
        #             #grava a req
        #             pyautogui.hotkey('Alt', 'G')
        #             time.sleep(0.5)
                    
        #             #Verifica se possui mensagem de produto incompativel
        #             funcoes_pyautogui.incompativel(produto_incompativel)
        #             time.sleep(1)
                    
        #             #Valida se a req está sem estoque
        #             sem_estoque = pyautogui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
        #             time.sleep(0.5)
                    
        #             #Verifica se aparece a tela de login de sem estoque
        #             funcoes_pyautogui.estoque(sem_estoque)
        #             time.sleep(0.5)
                    
        #             #fecha a tela
        #             time.sleep(1)
        #             pyautogui.moveTo(1220,730)
        #             pyautogui.click()
        #         break
        # # Caso a req seja do mesmo pedido, inclui uma nova sequencial
        # else:
        #     quebrar = None
        #     while quebrar == None:
        #         #Verifica se a primeira req teve erro de dinâmica
        #         if req_anterior == "Dinâmica":
        #             quebrar == "Dinâmica"
                
        #         #Clica no botão de Incluir Normal
        #         pyautogui.moveTo(450,130)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         #clica em "Sequencial - Via outra requisição"
        #         pyautogui.moveTo(245,370)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         #Clica no Botão OK
        #         pyautogui.moveTo(440,780)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         #Inclui o número da req
        #         pyautogui.moveTo(440,330)
        #         pyautogui.doubleClick()
        #         pyautogui.press('backspace')
        #         time.sleep(0.5)
        #         pyautogui.write(cod_item)
        #         pyautogui.press('enter')
        #         time.sleep(1)
                
        #         funcoes_pyautogui.tratamento(cod_item)
        #         time.sleep(0.5)
                
        #         #Confirma o item selecionado
        #         pyautogui.moveTo(1440,680)
        #         pyautogui.click()
        #         time.sleep(1)
                
        #         dinamica2 = pyautogui.locateOnScreen("imagens/dinamica2.png", confidence=0.9)
        #         if dinamica2 != None:
        #             pyautogui.press("enter")
        #             quebrar = "Dinâmica"
        #             req = "SEQ-Dinamica"
        #             break
                
        #         #Pula a tela de aviso
        #         pyautogui.moveTo(1440,720)
        #         pyautogui.click()            
        #         time.sleep(1)
                
        #         #Clica em editar Req
        #         pyautogui.moveTo(700,125)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         #ajusta a quantidade    
        #         pyautogui.moveTo(330,400)
        #         pyautogui.doubleClick()
        #         pyautogui.write(qtd)
        #         time.sleep(0.5)
                
        #         #calcula
        #         pyautogui.press('F9')
        #         time.sleep(1)
                
        #         produto_incompativel = pyautogui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
                
        #         #Pula a tela de produto incompatível
        #         pyautogui.press('enter')
        #         pyautogui.moveTo(x=950, y=660)
        #         pyautogui.click()
        #         time.sleep(1)
                
        #         #grava a req
        #         pyautogui.hotkey('Alt', 'G')
        #         time.sleep(0.5)
                
        #         funcoes_pyautogui.incompativel(produto_incompativel)
        #         time.sleep(0.5)
                
        #         #Valida se a req está sem estoque
        #         sem_estoque = pyautogui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
        #         time.sleep(0.5)
                
        #         funcoes_pyautogui.estoque(sem_estoque)
        #         time.sleep(0.5)
                
        #         #fecha a tela
        #         pyautogui.moveTo(1220,730)
        #         pyautogui.click()
        #         time.sleep(0.5)
                
        #         if cod_item in funcoes_pyautogui.kits:
        #             #Se o item for Tratamento para unhas ajusta a quantidade da outra sequencial
        #             pyautogui.moveTo(145,135)
        #             pyautogui.click()
        #             time.sleep(0.5)
                    
        #             #Clica em editar Req
        #             pyautogui.moveTo(700,125)
        #             pyautogui.click()
        #             time.sleep(0.5)
                    
        #             #ajusta a quantidade    
        #             pyautogui.moveTo(330,400)
        #             pyautogui.doubleClick()
        #             pyautogui.write(qtd)
        #             time.sleep(0.5)
                    
        #             #calcula
        #             pyautogui.press('F9')
        #             time.sleep(1)
                    
        #             #Verifica se possui mensagem de produto incompativel
        #             produto_incompativel = pyautogui.locateOnScreen('imagens/prod_incomp.png', confidence= 0.8)
        #             time.sleep(0.5)
                    
        #             #Pula a tela de produto incompatível e 
        #             pyautogui.press('enter')
        #             pyautogui.moveTo(x=950, y=660)
        #             pyautogui.click()
        #             time.sleep(1)
                    
        #             #grava a req
        #             pyautogui.hotkey('Alt', 'G')
        #             time.sleep(0.5)
                    
        #             #Verifica se possui mensagem de produto incompativel
        #             funcoes_pyautogui.incompativel(produto_incompativel)
        #             time.sleep(0.5)
                    
        #             #Valida se a req está sem estoque
        #             sem_estoque = pyautogui.locateOnScreen('imagens/sem-estoque.png', confidence= 0.8)
        #             time.sleep(0.5)
                    
        #             #Verifica se possui mensagem de sem estoque
        #             funcoes_pyautogui.estoque(sem_estoque)
        #             time.sleep(0.5)
                    
        #             #fecha a tela
        #             time.sleep(1)
        #             pyautogui.moveTo(1220,730)
        #             pyautogui.click()
        #             time.sleep(0.5)
        #         break
        # #Armazena o número do pedido como o anterior para verificação
        # num_pedido_ant = num_ped


    print('Tarefa concluída!') 
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
