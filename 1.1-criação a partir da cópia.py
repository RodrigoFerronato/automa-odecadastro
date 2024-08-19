import pyautogui
import time
import pandas
import pyperclip


pyautogui.PAUSE = 0.4

pyautogui.click(x=659, y=1059)

#Login no sistema
time.sleep(6)
pyautogui.click(x=940, y=538)
#MUDAR LOGIN CASO FOR OUTRA PESSOA
pyautogui.write("RODRIGO")
pyautogui.press("tab")
pyautogui.write("rodrigo")
pyautogui.press("enter")
pyautogui.click(x=57, y=94)
pyautogui.click(x=364, y=228)

#Lendo banco de dados
dtype_dict= {'chapa': str, 'fitac1': str,'fitac2': str,'fital1': str,'fital2': str}
tabela = pandas.read_excel("Criação cópia.xlsx", sheet_name="Filtragem", dtype=dtype_dict)
print(tabela)

#Entrando na tela de criação
pyautogui.press("C")
pyautogui.click(x=135, y=61)
pyautogui.press("backspace")

#Looping para preenchimento
for linha in tabela.index:

    novo = tabela.loc[linha, "novo"]
    pyautogui.write(novo)
    pyautogui.press("tab")
    try:
        # Verifica se a imagem está presente na tela com confiança ajustável
        if pyautogui.locateOnScreen('jacadastrado.png', confidence=0.9) is not None:
            pyautogui.press('enter')
            continue  # Pula para a próxima iteração do loop
    except pyautogui.ImageNotFoundException:
        # Se a imagem não for encontrada, continue normalmente
        pass
 
        
    pyautogui.write(tabela.loc[linha, "antigo"])
    pyautogui.press("tab", presses=5)

    pyautogui.write(str(tabela.loc[linha, "chapa"]))               
    pyautogui.press("tab", presses=7)

    pyautogui.write(tabela.loc[linha, "op"])
    pyautogui.press("tab", presses=12)
    pyautogui.doubleClick(x=73, y=307)
    pyautogui.rightClick(x=73, y=307)
    pyautogui.press('c')
    #Leitura do conteúdo do campo
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
    pyautogui.doubleClick(x=84, y=327)
    pyautogui.rightClick(x=84, y=327)
    pyautogui.press('c')
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
    pyautogui.rightClick(x=83, y=367)
    pyautogui.press('c')
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
    pyautogui.doubleClick(x=80, y=393)
    pyautogui.rightClick(x=80, y=393)
    pyautogui.press('c')
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
    pyautogui.rightClick(x=80, y=435)
    pyautogui.press('c')
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
    pyautogui.rightClick(x=82, y=477)
    pyautogui.press('c')
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
            pyautogui.press('tab', presses = 2)
    except ValueError:
        print("Texto não é um número inteiro no campo.")
    
    pyautogui.press("F2")
    pyautogui.press("tab", presses = 5)
    pyautogui.press("enter")





