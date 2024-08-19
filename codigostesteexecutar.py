import pyautogui
import time
import pandas
import pyperclip
import tkinter as tk
from tkinter import filedialog
import openpyxl

def executar_codigo(Modelo_de_estrutura_de_banco_de_dados):
    try:
        wb = openpyxl.load_workbook(Modelo_de_estrutura_de_banco_de_dados)
        pyautogui.PAUSE = 1

        pyautogui.click(x=659, y=1059)

        time.sleep(3)
        pyautogui.click(x=940, y=538)
        pyautogui.write("RODRIGO")
        pyautogui.press("tab")
        pyautogui.write("r")
        pyautogui.press("enter")
        pyautogui.click(x=57, y=94)
        pyautogui.click(x=364, y=228)

        dtype_dict= {'chapa': str, 'fitac1': str,'fitac2': str,'fital1': str,'fital2': str}
        tabela = pandas.read_excel("C:\PROJETOS\Código de automação trabalho manual\Criação de peças\Modelo de estrutura de banco de dados.xltx", sheet_name="Filtragem", dtype=dtype_dict)
        print(tabela)

        for linha in tabela.index:
            pyautogui.press("C")
            pyautogui.click(x=135, y=61)
            pyautogui.press("backspace")

            novo = tabela.loc[linha, "novo"]
            pyautogui.write(novo)
            pyautogui.press("tab")

            pyautogui.write(tabela.loc[linha, "antigo"])
            pyautogui.press("tab", presses=5)

            pyautogui.write(str(tabela.loc[linha, "chapa"]))               
            pyautogui.press("tab", presses=4)

            pyautogui.write(tabela.loc[linha, "uso"])
            pyautogui.press("tab", presses=3)

            pyautogui.write(tabela.loc[linha, "op"])
            pyautogui.press("tab", presses=12)
            time.sleep(0.5)
            pyautogui.doubleClick(x=73, y=307)
            time.sleep(0.5)
            pyautogui.rightClick(x=73, y=307)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab', presses = 2)
                    pyautogui.write(str(tabela.loc[linha, "fitac1"]))
                    pyautogui.press("tab")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
            time.sleep(0.5)
            pyautogui.doubleClick(x=84, y=327)
            time.sleep(0.5)
            pyautogui.rightClick(x=84, y=327)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab', presses = 2)
                    pyautogui.write(str(tabela.loc[linha, "fitac2"]))
                    pyautogui.press("tab")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
            # reconhecimento de campo termina aqui
            pyautogui.doubleClick(x=83, y=367)
            time.sleep(0.5)
            pyautogui.rightClick(x=83, y=367)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab', presses = 2)
                    pyautogui.write(str(tabela.loc[linha, "fital1"]))
                    pyautogui.press("tab")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
            # reconhecimento de campo termina aqui
            time.sleep(0.5)
            pyautogui.doubleClick(x=80, y=393)
            time.sleep(0.5)
            pyautogui.rightClick(x=80, y=393)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab', presses = 2)
                    pyautogui.write(str(tabela.loc[linha, "fital2"]))
                    pyautogui.press("tab")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
                
            pyautogui.doubleClick(x=80, y=435)
            time.sleep(0.5)
            pyautogui.rightClick(x=80, y=435)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab')
                    pyautogui.write(tabela.loc[linha, "parametro_externo"])
                    pyautogui.press("tab")
                    pyautogui.write("1")
                    pyautogui.press("tab")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
            pyautogui.doubleClick(x=82, y=477)
            time.sleep(0.5)
            pyautogui.rightClick(x=82, y=477)
            pyautogui.press('c')
            time.sleep(0.5)
            texto_no_campo = pyperclip.paste()
            try:
                valor = float(texto_no_campo.strip())
                if valor >= 1:
                    # Pressiona Tab duas vezes se o valor for maior que 1
                    pyautogui.press('tab')
                    pyautogui.write(tabela.loc[linha, "parametro_interno"])
                    pyautogui.press("tab")
                    pyautogui.write("1")
                else:
                    # Pressiona Tab três vezes se o valor for 1 ou menor
                    pyautogui.press('tab', presses = 3)
            except ValueError:
                print("Texto não é um número inteiro no campo.")
            
            pyautogui.press("F2")
            pyautogui.press("esc")
        wb.close()

        print(f"Código executado com sucesso usando o arquivo Excel: {Modelo_de_estrutura_de_banco_de_dados}")
    except Exception as e:
        print(f"Erro ao executar o código: {Modelo_de_estrutura_de_banco_de_dados}")

def on_executar_click():
    Modelo_de_estrutura_de_banco_de_dados = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
    if Modelo_de_estrutura_de_banco_de_dados:
        executar_codigo(Modelo_de_estrutura_de_banco_de_dados)

def on_fechar_click():
    root.destroy()

root = tk.Tk()
root.title("Seu Programa PyAutoGUI")

selecionar_arquivo_button = tk.Button(root, text="Selecionar Arquivo Excel", command=on_executar_click)
selecionar_arquivo_button.pack(pady=10)

fechar_button = tk.Button(root, text="Fechar", command=on_fechar_click)
fechar_button.pack(pady=5)

root.mainloop()





