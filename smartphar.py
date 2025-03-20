import os
import time
import pandas as pd
import pyautogui
import pyperclip
from dotenv import load_dotenv

from order import Order
from locations import *
from errors import *
from api_miliapp import atualizar_req_miliapp
import json

load_dotenv()
TOKEN_MILIAPP = os.getenv('TOKEN_MILIAPP')
LOGIN_SMARTPHAR = os.getenv('LOGIN_SMARTPHAR')
PASSWORD_SMARTPHAR = str(os.getenv('PASSWORD_SMARTPHAR'))
pyautogui.PAUSE = 0.3


def open_smartphar():
    print('open_smartphar')
    # Windows Run
    pyautogui.hotkey('win', 'r')
    pyautogui.write(r"\\10.1.1.5\j\smartpharmacy\SmartPhar.exe")
    pyautogui.press('enter')
    time.sleep(1.5)
    pyautogui.press(['left', 'enter'])
    time.sleep(3)


def login_smartphar(production_branch):
    print('login_smartphar')
    # Insert username
    pyautogui.write(LOGIN_SMARTPHAR)
    pyautogui.press('tab', presses=2)

    # Insert password
    pyautogui.write(PASSWORD_SMARTPHAR)
    pyautogui.press('tab')

    # Select branch
    if production_branch == '100':
        pyautogui.write
        pyautogui.press('down', presses=5)
        pyautogui.press('enter')
    elif production_branch == '600':
        pyautogui.write
        pyautogui.press(['down', 'enter'])
    
    time.sleep(1.3)
    pyautogui.press('enter', presses=2)
    time.sleep(1.3)
    pyautogui.press('enter')


def open_receitas_screen():
    print('open_receitas_screen')
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('r')
    time.sleep(0.1)
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('r')


def click_incluir_normal():
    print('click_incluir_normal')
    time.sleep(0.5)
    pyautogui.hotkey('alt', 'i')


def click_nova_receita_via_outra_requisicao():
    print('click_nova_receita_via_outra_requisicao')
    pyautogui.press('down', presses=3, interval=0.3)
    pyautogui.press("enter")


def click_sequencial_outra_requisicao():
    print('click_sequencial_outra_requisicao')
    pyautogui.press('down', presses=2, interval=0.3)
    pyautogui.press("enter")


def verificar_kit(sku):
    print('verificar_kit')
    kits_list = ['430035', '396864', '702310']
    if sku in kits_list:
        return True
    else:
        return False


def pesquisar_requisicao_inclusao_via_outra_receita(sku):
    print('pesquisar_requisicao_inclusao_via_outra_receita')
    time.sleep(1)
    try:
        pesquisa_requisicao = pyautogui.locateOnScreen('imagens/pesquisa_requisicao_3.PNG', confidence=0.9)
    except:
        pesquisa_requisicao = None
    print(f'pesquisa_requisicao: {pesquisa_requisicao}')

    if pesquisa_requisicao == None:
        pyautogui.press('tab', presses=4, interval=0.3)

    # pyautogui.moveTo(fied_receita_via_outra_requisicao_numero_req_location['x'], fied_receita_via_outra_requisicao_numero_req_location['y'])
    pyautogui.doubleClick()
    pyautogui.press('backspace')
    pyautogui.write(sku)
    time.sleep(0.3)
    pyautogui.press('enter', presses=2, interval=0.3)
    time.sleep(1)


def verify_max_dosage():
    print('verify_max_dosage')
    time.sleep(1.5)
    try:
        max_dosage = pyautogui.locateOnScreen('imagens/max_dosage.PNG', confidence=0.8)
    except Exception as e:
        print(e)
        max_dosage = None
    print(f'max_dosage: {max_dosage}')

    if max_dosage != None:
        pyautogui.press('enter')


def verify_min_dosage():
    print('verify_min_dosage')
    time.sleep(1.5)
    try:
        min_dosage = pyautogui.locateOnScreen('imagens/min_dosage.PNG', confidence=0.8)
    except Exception as e:
        print(e)
        min_dosage = None
    print(f'min_dosage: {min_dosage}')

    if min_dosage != None:
        pyautogui.press('enter')


def click_alterar():
    print('click_alterar')
    pyautogui.hotkey('alt', 'a')
    pyautogui.press('tab')


def transformar_or():
    print('transformar_or')
    pyautogui.write('OR')
    pyautogui.press('tab', presses=3)


def altera_data_hora_entrega(production_date):
    print('altera_data_hora_entrega')
    # Altera data
    pyautogui.write(production_date)
    pyautogui.press('tab')
    
    # Hora de entrega fix em 16:30
    pyautogui.write('163000')
    time.sleep(0.5)


def search_customer(cpf):
    print('search_customer')
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(cpf)
    pyautogui.press('enter')
    
    #verifica se cliente possui cadastro
    time.sleep(1)
    try:
        customer_verification = pyautogui.locateOnScreen('imagens/pesquisa_pessoas_ativas_2.PNG', confidence=0.8)
    except Exception as e:
        print(e)
        customer_verification = None
    print(f'customer_verification: {customer_verification}')

    if customer_verification != None:
        pyautogui.press('esc')
    else:
        pyautogui.press('enter')
    time.sleep(0.2)

    return customer_verification


def cadastrar_cliente(name, cpf, phone):
    print('cadastrar_cliente')
    pyautogui.press('insert')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write(name)
    pyautogui.press('tab', presses=4, interval=0.3)
    pyautogui.write(cpf)
    pyautogui.press('tab', presses=6, interval=0.3)
    pyautogui.write('M')
    pyautogui.press('tab', presses=9, interval=0.3)
    pyautogui.write(phone)
    pyautogui.hotkey('alt', 'g')


def incluir_crm(crm):
    print('incluir_crm')
    pyautogui.press('tab', presses=8, interval=0.3)
    pyautogui.write(crm)
    pyautogui.press('enter')
    time.sleep(1)


def ajustar_quantidade(qtd):
    print('ajustar_quantidade')
    pyautogui.press('tab', presses=12, interval=0.3)
    pyautogui.write(qtd)
    pyautogui.press('F9')
    time.sleep(0.5)
    pyautogui.press('enter', presses=3, interval=0.3)


def salva_req_o():
    print('salva_req_o')
    time.sleep(2)
    pyautogui.hotkey('alt', 'g')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    # Verifica tela de entrega
    time.sleep(1)
    try:
        print('locateOnScreen: entregar.PNG')
        entregar = pyautogui.locateOnScreen('imagens/entregar.PNG', confidence= 0.8)
    except:
        entregar = None
    print(entregar)

    if entregar != None:
        pyautogui.press(['right', 'enter'])

    # Verifica tela de imprimir orcamento
    time.sleep(1)
    try:
        print('locateOnScreen: imprimir_orcamento.PNG')
        imprimir_orcamento = pyautogui.locateOnScreen('imagens/imprimir_orcamento.PNG', confidence= 0.8)
    except:
        imprimir_orcamento = None
    print(imprimir_orcamento)

    if imprimir_orcamento != None:
        pyautogui.press(['right', 'enter'])


def save_req():
    print('save_req')
    pyautogui.hotkey('alt', 'g')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    # Verifica tela de sem estoque
    time.sleep(1)
    try:
        print('locateOnScreen: sem_estoque.PNG')
        error_estoque = verify_error_estoque()
    except:
        error_estoque = None
    print(error_estoque)

    # Verifica tela de entrega
    time.sleep(1)
    try:
        print('locateOnScreen: entregar.PNG')
        entregar = pyautogui.locateOnScreen('imagens/entregar.PNG', confidence= 0.8)
    except:
        entregar = None
    print(entregar)

    if entregar != None:
        pyautogui.press(['right', 'enter'])

    # Verifica tela de agenda compromissos
    time.sleep(1)
    try:
        print('locateOnScreen: agenda_compromissos.PNG')
        agenda_compromissos = pyautogui.locateOnScreen('imagens/agenda_compromissos.PNG', confidence= 0.8)
    except:
        agenda_compromissos = None
    print(agenda_compromissos)

    if agenda_compromissos != None:
        pyautogui.press('tab', presses=8, interval=0.3)
        pyautogui.press('enter')


def ajustar_quantidade_manual(qtd):
    print('ajustar_quantidade_manual')
    pyautogui.press('tab', presses=4, interval=0.3)
    pyautogui.write(qtd)
    pyautogui.press('F9')
    time.sleep(0.5)
    pyautogui.press('enter', presses=3, interval=0.3)


def get_req_number():
    print('get_req_number')
    req_number = None
    print('get_req_number')
    pyautogui.press('tab', presses=11, interval=0.3)
    pyautogui.hotkey('ctrl', 'c')
    req_number = pyperclip.paste()
    return req_number


def verify_orcamento_realizado():
    print('verify_orcamento_realizado')
    time.sleep(2)
    try:
        orcamento_realizado = pyautogui.locateOnScreen('imagens/orcamento_realizado.PNG', confidence= 0.8)
    except:
        orcamento_realizado = None
    print(f'orcamento_realizado: {orcamento_realizado}')

    if orcamento_realizado != None:
        pyautogui.press('enter')
        time.sleep(1)


def insert_orders_smartphar(smart_filtered_orders, sector_var, production_branch, production_date):
    print('insert_orders_smartphar')
    # Variables to receive order number and reqs
    previous_order_number = None
    error = None
    reqs = []
    reqs_list = []
    orders_errors = []
    
    for i in range(len(smart_filtered_orders)):
        print(f'===== {i+1} / {len(smart_filtered_orders)} =====')

        #Define uma variável para cada valor a ser usado no código
        sku = smart_filtered_orders["Código (SKU)"].iloc[i]
        name = smart_filtered_orders["Nome do contato"].iloc[i]
        cpf = smart_filtered_orders["CPF/CNPJ"].iloc[i]
        order_number = smart_filtered_orders["Número do pedido"].iloc[i]
        qtd = smart_filtered_orders["Quantidade"].iloc[i]
        phone = smart_filtered_orders["Fone"].iloc[i]
        req_number = None
        
        order = Order(sku, name, cpf, order_number, qtd, phone)
        print(order)

        if error == 'dinamica':
            if order.order_number == previous_order_number:
                continue
            else:
                error = None

        print(f'previous_order_number: {previous_order_number}')
        print(f'order.order_number == previous_order_number: {order.order_number == previous_order_number}')

        if order.order_number != previous_order_number:
            while True:
                try:
                    # Clica no botão de Incluir Normal
                    time.sleep(1)
                    click_incluir_normal()
                    time.sleep(1.5)

                    # Verifica se está na tela correta
                    tela_incluir = pyautogui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)
                    print(f'tela_incluir: {tela_incluir}')
                except Exception as e:
                    print('Erro tela_incluir !=')
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
            # Clica no botão de Incluir Normal
            time.sleep(1)
            click_incluir_normal()
            time.sleep(1.5)

            try:
                # Verifica se está na tela correta
                tela_incluir = pyautogui.locateOnScreen("imagens/tela_incluir.png", confidence=0.8)
                print(f'tela_incluir: {tela_incluir}')
            except:
                print('Erro tela_incluir ==')
                tela_incluir = None
                continue

            click_sequencial_outra_requisicao()
            time.sleep(1)

        # Clica em Nova Receita via outra requisição
        print(f'incluindo uma sequencial pedido: {order.order_number}: {order.sku}')
        pesquisar_requisicao_inclusao_via_outra_receita(order.sku)
        time.sleep(1.5)
        previous_order_number = order.order_number
        print('previous_order_number = order.order_number')
        
        #Verifica se possui erro de dinamica
        error_dinamica = verify_error_dinamica()
        time.sleep(0.5)

        if error_dinamica != None:
            error = 'dinamica'
            e = new_error(error, order.order_number)
            orders_errors.append(e)

            # Fecha a tela de erro
            pyautogui.press('esc')
            continue

        error_estoque = verify_error_estoque()
        time.sleep(0.5)

        if error_estoque != None:
            error = 'estoque'
            e = new_error(error, order.order_number)
            orders_errors.append(e)

            # Fecha a tela de erro
            pyautogui.press('esc')
            continue
        
        # Confirma item
        print('confirmando item')
        pyautogui.press('enter')
        time.sleep(0.5)

        # Verifica dosagem máxima / mínima
        verify_max_dosage()
        verify_min_dosage()

        # Verifica tela de orçamento realizado
        verify_orcamento_realizado()

        # Alterar req - quantidade
        print(f'sector_var: {sector_var}')
        print(f'qtd: {qtd}')
        if sector_var == 'Manual':
            click_alterar()
            
            # Incluir CRM
            crm = smart_filtered_orders["Observações"].iloc[i]
            incluir_crm(str(crm))

            if int(qtd) > 1:
                ajustar_quantidade_manual(order.qtd)
            salva_req_o()
        else:
            if int(qtd) > 1:
                click_alterar()
                ajustar_quantidade(order.qtd)
                salva_req_o()
        print('Sequencial incluída em O com sucesso!')

        try:
            next_order_number = smart_filtered_orders["Número do pedido"].iloc[i+1]
            print(f'order.order_number != str(next_order_number) {order.order_number != str(next_order_number)}')
        except:
            next_order_number = None
        print(f'next_order_number: {next_order_number}')

        if order.order_number != str(next_order_number):
            click_alterar()
            transformar_or()
            altera_data_hora_entrega(production_date)

            pyautogui.press('tab', presses=2, interval=0.3)
            # pyautogui.press('tab')
            # pyautogui.press('tab')

            customer_verification = search_customer(order.cpf)
            if customer_verification != None:
                cadastrar_cliente(order.name, order.cpf, order.phone)
            save_req()

            req_number = get_req_number()
            print(req_number)

            if req_number.isdigit():
                print('is digit')
            else:
                print('not digit')
                req_number = None
            
            if req_number not in reqs_list and req_number != None and req_number != '':
                req_obj = {
                    'numero_tiny': order.order_number,
                    'req_smart': req_number,
                    'cnpj': '07413904000198'
                }
                print(req_obj)
                reqs_list.append(req_number)
                reqs.append(req_obj)
                atualizar_req_miliapp(TOKEN_MILIAPP, req_obj)
            else:
                new_error('Req não é número', order.order_number)

            with open('errors.json', 'w') as errors_file:
                json.dump(orders_errors, errors_file, indent=4)

            # Save reqs to a JSON file
            with open('reqs.json', 'w') as reqs_file:
                json.dump(reqs, reqs_file, indent=4)

    print('Tarefa concluída!') 
    # Save errors to a JSON file
    with open('errors.json', 'w') as errors_file:
        json.dump(orders_errors, errors_file, indent=4)

    # Save reqs to a JSON file
    with open('reqs.json', 'w') as reqs_file:
        json.dump(reqs, reqs_file, indent=4)