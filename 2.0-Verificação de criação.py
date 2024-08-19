import pyautogui
import time
import pandas
import pyperclip

pyautogui.PAUSE = 0.8
#Leitura do banco de dados
dtype_dict= {'chapa': str, 'fitac1': str,'fitac2': str,'fital1': str,'fital2': str}
tabela = pandas.read_excel("Criação cópia.xlsx", sheet_name="Filtragem", dtype=dtype_dict)
print(tabela)

pyautogui.click(x=132, y=58)
#Digitando o código para confirmar o cadastro
for linha in tabela.index:
    novo = tabela.loc[linha, "novo"]
    pyautogui.write(novo)
    pyautogui.press("enter", presses = 2)

