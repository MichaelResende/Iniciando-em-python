#!/usr/bin/env python
# coding: utf-8

# In[38]:


#---------------------------------------------------------------------------------------------------------------#
#                                               Importações                                                     #
#---------------------------------------------------------------------------------------------------------------#

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options #Abre o navegador em segundo plano
from time import sleep
import pandas as pd
import urllib
import urllib.request
import getpass
from tkinter import *
from tkinter import messagebox


#---------------------------------------------------------------------------------------------------------------#
#                                               Funções                                                         #
#---------------------------------------------------------------------------------------------------------------#

def linha(tam = 100):
    return '-' * tam

def cabecalho(txt):
    print(linha())
    print(txt.center(100))
    print(linha())

#---------------------------------------------------------------------------------------------------------------#
#                                               Listas                                                          #
#---------------------------------------------------------------------------------------------------------------#

CR = list()

#---------------------------------------------------------------------------------------------------------------#
#                                               Programa                                                        #
#---------------------------------------------------------------------------------------------------------------#

#Abrir navegador em segundo plano
chrome_options = Options()
chrome_options.headless = True #Sem aparecer

try:
    cabecalho('Sempre utilize o PROMPT, ignore a aba do Chrome!')
    cabecalho('Estamos acessando o site, aguarde!')

    #Abre o navegador
    navegador = webdriver.Chrome()
    #navegador = webdriver.Chrome(options=chrome_options)
    navegador.get('https://redmine.algartech.com')

    while True: 
        #Login
        navegador.find_element_by_xpath('//*[@id="username"]').clear()
        sleep(5)
        print()
        cabecalho('Informe o login e senha do redmine.')
        navegador.find_element_by_xpath('//*[@id="username"]').send_keys(input('Login: ').split())

        #Senha
        navegador.find_element_by_xpath('//*[@id="password"]').clear()
        navegador.find_element_by_xpath('//*[@id="password"]').send_keys(getpass.getpass('Senha: '))
        navegador.find_element_by_xpath('//*[@id="login-form"]/form/table/tbody/tr[4]/td[2]/input').send_keys(Keys.ENTER)
        print(linha())

        #Projetos
        try:
            sleep(1)
            navegador.refresh()
            navegador.find_element_by_xpath('//*[@id="top-menu"]/ul/li[2]/a').click()
        except:
            cabecalho('ERRO: Login ou senha invalidos. Caso contrario contate o desenvolvedor (Michael Resende)!')
            sleep(2)
            navegador.refresh()
            continue
        else:
            break

    #01 - BI - Business Intelligence - RELATORIOS
    navegador.find_element_by_xpath('//*[@id="projects-index"]/ul/li[1]/div/a').click()

    #Tempo Gasto
    navegador.find_element_by_xpath('//*[@id="sidebar"]/p[2]/a[1]').click()

    try:
        tb_Demanda = pd.read_excel('Demandas.xlsx')
        cont = tb_Demanda['Tarefa'].count()
    except:
        print('ERRO: Arquivo "Demandas.xlsx" não localizado ou no padrao incorreto...')

    try:
        tb_CRs = pd.read_excel('Operação.xlsx')
        con = tb_CRs['Operação'].count()
    except:
        print('ERRO: Arquivo "Operação.xlsx" não localizado ou no padrao incorreto...')
    else:
        for i in range(0, con):
            CR.append(tb_CRs['Operação'][i])

    cabecalho('Lançando demandas, aguarde!')

    for l in range(0, cont):
    #Tarefa
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_issue_id"]').clear()
            navegador.find_element_by_xpath('//*[@id="time_entry_issue_id"]').send_keys(int(tb_Demanda['Tarefa'][l]))
        except:
            print(f'ERRO: A TAREFA da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Data
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_spent_on"]').clear()
            navegador.find_element_by_xpath('//*[@id="time_entry_spent_on"]').send_keys(str(tb_Demanda["Data"][l])[:10])
        except:
            print(f'ERRO: A DATA da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Horas
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_hours"]').clear()
            navegador.find_element_by_xpath('//*[@id="time_entry_hours"]').send_keys(str(tb_Demanda["Horas"][l])[:5])
        except:
            print(f'ERRO: O HORARIO da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Comentario
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_comments"]').clear()
            navegador.find_element_by_xpath('//*[@id="time_entry_comments"]').send_keys(tb_Demanda["Comentario"][l])
        except:
            print(f'ERRO: O COMENTARIO da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Atividade
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_activity_id"]').click()
            navegador.find_element_by_xpath('//*[@id="time_entry_activity_id"]').send_keys(tb_Demanda["Atividade"][l])
        except:
            print(f'ERRO: A ATIVIDADE da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Operacao
        try:
            for o in CR:
                navegador.find_element_by_xpath(o).click()
        except:
            print(f'ERRO: Os CRs informados estão incorretos, tente novamente...')
            continue
    #Cliente
        try:
            navegador.find_element_by_xpath('//*[@id="time_entry_custom_field_values_24"]').click()
            navegador.find_element_by_xpath('//*[@id="time_entry_custom_field_values_24"]').send_keys(tb_Demanda["Cliente"][l])
        except:
            print(f'ERRO: O CLIENTE da linha {l} está com erro, valide o arquivo e tente novamente...')
            continue
    #Criar
        try:
            navegador.find_element_by_xpath('//*[@id="new_time_entry"]/input[5]').click()
            sleep(2)
            navegador.refresh()
        except:
            print(f'ERRO: O botão de lançamento foi alterado, contate o desenvolvedor (Michael Resende)')
            continue

    navegador.quit()

    janela = Tk()
    messagebox.showinfo('Concluído', 'Lançamento realizado, valide as demanas no redmine.')
    janela.destroy()

except:
    cabecalho('ERRO: Algo de inesperado aconteceu... Tente novamente!')

