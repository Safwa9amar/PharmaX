from ast import Str
import pandas as pd
import pyautogui as astro
import pyperclip
import time
from tqdm import tqdm
import cleantext


def getDataDict(df, arr):
    df.columns = arr
    df.head()
    data = df.to_dict()
    data_dict = []
    for index, row in df[df.columns].iterrows():
        data_dict.append({row.to_list()[0]: row.to_list()[1:]})
    return data_dict
print('============================================')

sheet = input('Sheet Number :')
lot1 = pd.read_excel('04.xlsx',f'Sheet{sheet}')
lot1Dict = getDataDict(lot1, ['name', 'qty', 'pu', '%R',
                       'Prix AR', 'TVA', 'PPA', 'N°Lot', 'SHP', 'Peremp', ''])


def getDrogIdx(index):
    new = input(f'Prev:{index} / Next : {index} => New : ')
    if new == '':
        return index
    else:
        return new

index = 0
while int(index) < len(lot1Dict):
    index = int(getDrogIdx(index))
    dic = lot1Dict[int(index)]
    print(dic)
    for key in dic.items():
        index = index + 1
        file = open('counter2.txt', 'a')
        file.write('\n' + str(index))
        file.close()
        astro.moveTo(x=1046, y=-150)
        astro.click()
        pyperclip.copy(cleantext.clean(key[0], numbers=False, punct=True))
        astro.hotkey('ctrl', 'v')


        if input('Continue.........') == '':
            astro.moveTo(1654, 39)
            astro.click()
            pyperclip.copy(key[1][6])
            astro.hotkey('ctrl', 'v')
        if input('Continue.........') == '':
            astro.moveTo(x=1560, y=97)
            astro.doubleClick()
            time.sleep(0.3)
            astro.write(str(key[1][0]))
        if int(key[1][2]) == 100:
            astro.moveTo(x=1521, y=120)
            astro.doubleClick()
            time.sleep(0.3)
            astro.write(str(key[1][1]))

        print('============================================')
