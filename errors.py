import pyautogui


def new_error(error, order_number):
    print(f'ERROR: {error}!')
    new_error = {
        'order_number': order_number,
        'error': error
    }
    return new_error


def verify_error_dinamica():
    print('verify_error_dinamica')
    try:
        error_dinamica = pyautogui.locateOnScreen('imagens/erro_dinamica.PNG', confidence=0.8)
    except:
        error_dinamica = None
    
    return error_dinamica


def verify_error_estoque():
    print('verify_error_estoque')
    try:
        error_estoque = pyautogui.locateOnScreen('imagens/sem_estoque.PNG', confidence= 0.8)
        
        pyautogui.write("MILIBOT")
        for _ in range(2):
            pyautogui.press('tab')
        pyautogui.write("77777777")
        pyautogui.press('enter')
    except:
        error_estoque = None
    
    return error_estoque

