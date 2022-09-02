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


lot1 = pd.read_excel('lot2.xlsx')
lot1Dict = getDataDict(lot1, ['name', 'qty', 'pu', '%R',
                       'Prix AR', 'TVA', 'PPA', 'N°Lot', 'SHP', 'Peremp', ''])


def getDrogIdx(index):
    new = input(f'Prev:{index} / Next : {index} => New : ')
    if new == '':
        return index
    else:
        return new

print('============================================')
index = 0
while int(index) < len(lot1Dict):
    index = int(getDrogIdx(index))
    dic = lot1Dict[int(index)]
    print(pd.DataFrame(dic))
    for key in dic.items():
        index = index + 1
        file = open('counter.txt', 'a')
        file.write('\n' + str(index))
        file.close()
        astro.moveTo(1075, -94)
        astro.click()
        pyperclip.copy(cleantext.clean(key[0], numbers=False, punct=True))
        astro.hotkey('ctrl', 'v')
        time.sleep(0.5)
        PU = str(key[1][1])
        R = str(key[1][2])
        PPA = str(key[1][5])
        SHP = str(key[1][7])
        time.sleep(0.5)
        # get data from day obj
        if type(key[1][8]) == str:
            day = str(key[1][8])[:1]
            month = str(key[1][8])[3:5]
            year = str(key[1][8])[6:]
        else:
            day = str(key[1][8].day)
            month = str(key[1][8].month)
            year = str(key[1][8].year)

        
        if input('Continue.........') == '':
            astro.moveTo(1634, -12)
            astro.click()
            pyperclip.copy(key[1][6])
            astro.hotkey('ctrl', 'v')

            astro.press('tab')


            astro.typewrite(day)
            astro.press('right')
            astro.typewrite(month)
            astro.press('right')
            astro.typewrite(year)

            astro.press('tab')

            print((PU), (R), (PPA), (SHP))

            astro.write(PU)
            astro.press('tab')

            astro.write(R)
            astro.press('tab')

            astro.write(PPA)
            astro.press('tab')

            astro.write(SHP)
            astro.press('tab')
        print('============================================')
