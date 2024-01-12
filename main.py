
import pyautogui
import pytesseract
import time

import win32gui


pytesseract.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"
def pegar_regiao(img,plus_x=0,plus_y=0,width=0,height=0):

    box=pyautogui.locateOnScreen(img,grayscale=True,confidence=0.8)
    x,y= box[0],box[1]
    x=int(x+plus_x)
    y=int(y+plus_y)
    width =int(width)
    height =int(height)
    return pyautogui.screenshot(region=(x,y,width,height))





ang=pegar_regiao('img/angulo.png',plus_x=80,width=95,height=45)
dist=pegar_regiao('img/dist.png',plus_x=80,width=100,height=25)
alt=pegar_regiao('img/altura.png',plus_x=80,width=70,height=25)
vento=pegar_regiao('img/vento.png',plus_x=80,width=55,height=25)
terre=pegar_regiao('img/terreno.png',plus_x=80,plus_y=-5,width=70,height=25)
lista=[ang,dist,alt,vento,terre]
lista_valores=[]
for i in lista:

    texto=pytesseract.image_to_string(i,config='--psm 6')
    texto=texto.replace('%','').replace('m','').replace('n','').replace("\n",'')


    lista_valores.append(texto)

print(lista_valores)
lista_coordenadas=[
    {
        'nome': 'angulo',
        'valor': lista_valores[0],
        'coordenadas': (-1240,216)
    },
    {
        'nome': 'distancia',
        'valor': lista_valores[1],
        'coordenadas': (-1223,138)
    },
    {
        'nome': 'altura',
        'valor': lista_valores[2],
        'coordenadas': (-1224,165)
    },
    {
        'nome': 'vento',
        'valor': lista_valores[3],
        'coordenadas': (-1224,191)
    },
    {
        'nome': 'terreno',
        'valor': lista_valores[4],
        'coordenadas': (-1225,295)
    }
]

hwnd = win32gui.FindWindow(None, '290+10tibas - Excel')

if hwnd != 0:
    win32gui.SetForegroundWindow(hwnd)

for valor in lista_coordenadas:
    pyautogui.click(valor['coordenadas'])
    pyautogui.write(valor['valor'])
pyautogui.press('enter')
