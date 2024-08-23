import xlwings as xw
import keyboard as kb
import time
import random


#change column
#pos = pos[:0] + 'B' + pos[1:]
#change line
#pos = pos[:1] + '2' + pos[2:]


wb = xw.Book('Example.xlsx')
sht1 = wb.sheets['Sheet']
corpo = ''
pos = [10,15]
prevPos = pos
size = 2
indoRight = (1,1)


##print('go')


def spawnfruit():
    fruitX = random.randrange(10,45)
    fruitY = random.randrange(10,18)
    
    fruitlocal = (fruitX,fruitY)
    print(fruitlocal)
    
    sht1.range(fruitlocal).color = (255,0,0)

spawnfruit()

def collison(pos,size):
    if sht1.range(pos).color == (255,0,0):
        size = size + 1
        return size
    else:
        pass
        



while True:
    pos = list(pos)
    prevPos = list(prevPos)
    if kb.is_pressed('W'):
        
        pos = list(pos)
        prevPos = list(prevPos)
        
        if pos[0] - size <= 0 or pos[1] <= 0:
            prevPos[1] = pos[1]
            prevPos[0] = pos[0]
        else:
            prevPos[0] = pos[0] + size
            prevPos[1] = pos[1] 
        
        pos[0] = pos[0] - 1
        
        
        
        pos = tuple(pos)
        prevPos = tuple(prevPos)
        
        

        
        if pos[1] - size <= 0:
            tutu2 = (pos[0], pos[1])
        elif pos[0] - size <= 0:
            tutu2 = (pos[0] + size, pos[1])
        elif indoRight[1] <= pos[1]:
             tutu2 = (pos[0] + size, pos[1] - size)
        else:
            tutu2 = (pos[0] + size, pos[1] + size)
             
        if sht1.range(pos).color == (255,0,0):
            size = size + 1
            spawnfruit()
        elif sht1.range(tutu2).color == (255,0,0):
            spawnfruit()
        else:
            pass
        
        sht1.range(prevPos,tutu2).color = (255,255,255)

        sht1.range(prevPos).color = (255,255,255)
        
        sht1.range(pos).color = (0, 0, 0)
        time.sleep(0.1)
        
        print('Posição atual: ', pos)
        print('posição previa: ', prevPos)
        print('indo pra direita: ', indoRight)
        print('tutu2: ', tutu2)
        print('tamanho: ', size)
        
        
        
        
        # ('Vertical previo', prevPos)       
        
        
        
    if kb.is_pressed('A'): 
        pos = list(pos)
        prevPos = list(prevPos)
        
        
        if pos[1] - size <=0 or prevPos[1] <= 1:
            prevPos[0] = pos[0]
            prevPos[1] = pos[1] + size
        else:
            prevPos[0] = pos[0]
            prevPos[1] = pos[1] + size
        
        if pos[1] <= 1:
            pos[1] = 1
        else:
            pos[1] = pos[1] - 1
        # ('OI: ', pos)
        
        
    
        pos = tuple(pos)
        prevPos = tuple(prevPos)
        
        if sht1.range(pos).color == (255,0,0):
            size = size + 1
            spawnfruit()
        elif sht1.range(tutu2).color == (255,0,0):
            spawnfruit()
        else:
            pass
        
        
        if pos[1] - size <= 0:
            tutu2 = (pos[0], pos[1])
        elif pos[0] - size <= 0:
            tutu2 = (pos[0], pos[1] + size)
        else:
            tutu2 = (pos[0] - size, pos[1] + size)
        
        sht1.range(prevPos,tutu2).color = (255,255,255)
        
        sht1.range(pos).color = (0,0,0)
        
        sht1.range(prevPos).color = (255,255,255)
        
        time.sleep(0.1)
        
    if kb.is_pressed('D'):
        pos = list(pos)
        prevPos = list(prevPos)
        
        if prevPos[1] - 1 == 0 or pos[1] - size <= 0 or pos[0] - size <= 0:
            prevPos[0] = pos[0]
            prevPos[1] = pos[1] + size
        else:
            prevPos[0] = pos[0] 
            prevPos[1] = pos[1] - size
            
            
        pos[1] = pos[1] + 1
        
        
        pos = tuple(pos)
        prevPos = tuple(prevPos)
        indoRight = pos
        
        
        ###print(indoRight)
        
        
        if pos[1] - size <= 0:
            tutu2 = (pos[0], pos[1])
        elif pos[0] - size <= 0:
            tutu2 = (pos[0], pos[1] - size)
        else:
            tutu2 = (pos[0] - size, pos[1] - size)
        

        if sht1.range(pos).color == (255,0,0):
            size = size + 1
            spawnfruit()
        elif sht1.range(tutu2).color == (255,0,0):
            spawnfruit()
        else:
            pass
        
        sht1.range(prevPos,tutu2).color = (255,255,255)
        
        print('Posição atual: ', pos)
        print('posição previa: ', prevPos)
        print('indo pra direita: ', indoRight)
        print('tutu2: ', tutu2)
        print('tamanho: ', size)
        
        sht1.range(pos).color = (0, 0, 0)
        sht1.range(prevPos).color = (255,255,255)
        
        time.sleep(0.1)
        
        # ###print("Horizontal Atual: ",pos)
        ####print("Horizontal Anterior: ", prevPos)
        
    if kb.is_pressed('S'):
        pos = list(pos)
        prevPos = list(prevPos)
        
        if pos[0] - size <= 0 or pos[1] <= 0:
            prevPos[1] = pos[1]
            prevPos[0] = pos[0]
        else:
            prevPos[0] = pos[0] - size
            prevPos[1] = pos[1]
        
        pos[0] = pos[0] + 1
        
        pos = tuple(pos)
        prevPos = tuple(prevPos)
        
        
        
        
        
        
        ##print("S: ",indoRight[1])
        
        if pos[1] - size <= 0:
            tutu2 = (pos[0], pos[1])
        elif pos[0] - size <= 0:
            tutu2 = (pos[0] + size, pos[1])
        elif indoRight[1] <= pos[1]:
             tutu2 = (pos[0] - size, pos[1] - size)
        else:
            tutu2 = (pos[0] - size, pos[1] + size)
        
        sht1.range(prevPos,tutu2).color = (255,255,255)

        if sht1.range(pos).color == (255,0,0):
            size = size + 1
            spawnfruit()
        elif sht1.range(tutu2).color == (255,0,0):
            spawnfruit()
        else:
            pass
        
        sht1.range(prevPos).color = (255,255,255)
        
        sht1.range(pos).color = (0, 0, 0)
        
        time.sleep(0.1)
        
        ###print('Vertical previo', prevPos)

    
    
        
