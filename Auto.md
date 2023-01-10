#Passo 1: Entrar no sistema da empresa -- site.....
pyautogui.PAUSE = 1

time.sleep(2)
pyautogui.hotkey("ctrl", "t") #nova aba
pyperclip.copy("local onde o arquivo está alocado para download")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

#Passo 2: Navegar até o local do relatório -- pasta exportar
pyautogui.click(x=417, y=267, clicks=2)
time.sleep(4)

#Passo 3: Exportar o relatório -- fazer download
pyautogui.click(x=404, y=328)
pyautogui.click(x=1721, y=151)
time.sleep(2)
pyautogui.click(x=1536, y=569)
time.sleep(5)

#Passo 4: Calcular os indicadores -- faturamento e quantidade de produtos
import pandas

tabela = pandas.read_excel(r"pasta onde o arquivo esta após download")
display(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

#Passo 5: Enviar um e-mail para diretoria 
#abrir aba e entrar e-mail
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("e-mail por onde será enviado os dados")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(8)

#clicar botão escrever
pyautogui.click(x=89, y=173)
time.sleep(5)
#preencher as informações do e-mail
#destinatário
pyautogui.write("e-mail do destinatário")
pyautogui.press("tab") #selecionar email

#assunto
#pular para campo assunto
time.sleep(2)
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

#corpo
texto = f"""Prezados, bom dia

O faturamento de ontem foi de : R${faturamento:,.2f}
A quantidade vendida foi de: {quantidade:,}

Att."""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)
#enviar o e-mail
pyautogui.hotkey("ctrl", "enter")
