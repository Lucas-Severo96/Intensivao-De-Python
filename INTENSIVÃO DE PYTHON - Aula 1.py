#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 1
pyautogui.hotkey("ctrl", "t")
pyautogui.alert("Vai começar, aperte OK e não mexa em nada!")


# In[2]:


link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(468, 464, clicks=2)
pyautogui.click(467, 611)
pyautogui.click(button='right')
pyautogui.click(629, 919)


# In[3]:


# Janela para Salvar o arquivo!
time.sleep(5)
pyautogui.click(1467, 963)


# In[4]:


df = pd.read_excel(
    r"C:\Users\nc\Desktop\MATERIAL PROGRAMAÇÃO\INTENSIVÃO DE PYTHON\Vendas - Dez.xlsx")
display(df)


# In[5]:


# Cálculo dos Indicadores
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()


# In[6]:


# Abrir nova aba do Gmail
pyautogui.hotkey("ctrl", "t")
pyautogui.write("mail.google.com")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(107, 287)


# In[7]:


# Escrevendo e-mail
pyautogui.write("lucasleite.lls+diretoria@gmail.com")
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")


# In[8]:


# Escopo do e-mail
pyautogui.press('tab')
texto = f"""
Prezados, Bom Dia!

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos vendidos foi {qtde_produtos:,}

Abs
Lucas Severo"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
