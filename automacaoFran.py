import pyautogui
import time


#abre o google
pyautogui.click(x=592, y=884)
time.sleep(3)

#abre o tine
pyautogui.write("http://tiny.com.br/login")
pyautogui.press("Enter")

time.sleep(3)

#abre o norton e coloca a senha do cofre para o ling
pyautogui.click(x=943, y=361)
time.sleep(3)
pyautogui.click(x=766, y=559)
time.sleep(3)
pyautogui.write("20141002304220Ksl*")
time.sleep(3)
pyautogui.press("Enter")
time.sleep(6)

#loga no tiny
pyautogui.click(x=943, y=361)
time.sleep(3)
pyautogui.click(x=824, y=496)
time.sleep(3)
pyautogui.click(x=814, y=519)
time.sleep(8)

#abre as finanças
pyautogui.moveTo(x=27, y=277, duration=2.5)
time.sleep(3)
pyautogui.click(x=137, y=443)
time.sleep(2)
pyautogui.click(x=322, y=275)
time.sleep(3)
pyautogui.click(x=1045, y=226)
time.sleep(4)
pyautogui.click(x=867, y=220)
time.sleep(4)
pyautogui.click(x=970, y=387)
time.sleep(3)
pyautogui.click(x=1087, y=437)
time.sleep(3)

#seleciona o intervalo Jan - Atual
pyautogui.click(x=834, y=478, clicks=5)
time.sleep(3)
pyautogui.click(x=875, y=538)
time.sleep(3)
pyautogui.click(x=815, y=543)
time.sleep(3)

#tira o filtro de conta
pyautogui.click(x=373, y=279)
time.sleep(3)
pyautogui.click(x=379, y=363)
