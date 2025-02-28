import pyautogui
import time
 
def get_mouse_position():
    x, y = pyautogui.position()
    return x, y
 
def move_mouse_to_position(x, y):
    pyautogui.moveTo(x, y, duration=1)

time.sleep(5)
position = get_mouse_position()
move_mouse_to_position(position[0], position[1])
print(position)