#!/usr/bin/env python
# coding: utf-8

# ###                                                                 bibliotecas

# In[34]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options #Abre o navegador em segundo plano
from time import sleep
import pandas as pd
from tkinter import *
from tkinter import ttk #Barra de progresso
import urllib.request #Testa o site


# #### Validação de arquivo

# In[35]:


def arquivo():

    texto('Olá, seja bem vindo!',0)
    
    try:
        tb_Demanda = pd.read_excel('Demandas.xlsx')
        cont = tb_Demanda['Tarefa'].count()
    except:
        texto('ERRO: Arquivo "Demandas.xlsx" não localizado ou no padrao incorreto.', 1)
        b_arq.grid(row=2, column=1, pady=5, sticky=W)
        
    if cont == 0:
        texto('Identificamos que o arquivo de demandas está vazio, clique em "Buscar arquivo novamente" para atualizar', 1)
        b_arq.grid(row=2, column=1, pady=5, sticky=W)
    else:
        texto_branco(1)
        texto(f'Identificamos {cont} tarefas para lançamento.', 1)
        texto_linha(2)
        texto('Informe o login e senha do redmine.', 3)
        t_login.grid(row=4, column=0, pady=5, sticky=E)
        t_senha.grid(row=5, column=0, pady=5, sticky=E)
        c_login.grid(row=4, column=1, pady=5, sticky=W)
        c_senha.grid(row=5, column=1, pady=5, sticky=W)
        b_login.grid(row=6, column=1, pady=5, sticky=W)
        texto_linha(8)
        


# ### Validação login e senha

# In[36]:


def login():
    
#Valida se o Chrome Drive está atualizado
    try:
        navegador = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    except:
        texto('ERRO: CHROMEDRIVE atualizou! Pesquise CHROMEDRIVE no Google e baixe a nova versão. Dúvidas contate o desenvolvedor (Michael Resende)!', 7)
        texto_linha(8)    
    else:
#Valida se e possivel acessar o site
        try:
            site = urllib.request.urlopen('http://redminealgar/login')
                                         #'https://redmine.algartech.com/login'
        except:
            texto_branco(7)
            texto('ERRO: O site do REDMINE está indisponivel, verifique a internet ou proxy!. Caso contrario contate o desenvolvedor (Michael Resende)!', 7)
            texto_linha(8)
        else:
            login = str(c_login.get())
            senha = str(c_senha.get())

            try:
#Abre o navegador
                #chrome_opt = webdriver.ChromeOptions()
                #chrome_opt.add_argument('headless')
                #navegador = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_opt)
                navegador.get('http://redminealgar/login')
                             #'https://redmine.algartech.com/login')
                navegador.find_element_by_id('username').send_keys(login)
                navegador.find_element_by_id('password').send_keys(senha)
                navegador.find_element_by_name('login').send_keys(Keys.ENTER)
            
                ##navegador.find_element_by_xpath('//*[@id="top-menu"]/ul/li[2]/a').click()
                navegador.find_element_by_class_name('projects').click()
                navegador.quit()
                #Exibe comboBox
                texto_branco(7)
                texto_branco(3)
                texto('Login realizado com sucesso! Selcione sua área:', 4)
                texto_branco(5)
                areas = ['BI', 'PCP', 'Trn / Qld']
                cb_areas = ttk.Combobox(app, values=areas)
                cb_areas.set('BI')
                cb_areas.grid(row=5, column=1, pady=5, sticky=W)
                bt_area = Button(app, text='Confirme a área',command = lambda: validaArea(cb_areas.get()))
                bt_area.grid(row=6, column=1, pady=5, sticky=W)

            except:
                texto_branco(7)
                texto('ERRO: Login ou senha invalidos. Caso contrario contate o desenvolvedor (Michael Resende)!', 7)
                texto_linha(8)
                sleep(2)
                navegador.quit()


# ### Valida a área do usuário

# In[37]:


def validaArea(vArea):
    if vArea == 'PCP':
        texto(f'Área selecionada foi {vArea}. Clique em "Lançar demandas" para prosseguir.', 7)
        LinkArea = 'http://redminealgar/projects/gep_pcp/time_entries/new'
                  #'https://redmine.algartech.com/projects/gep_pcp/time_entries/new'
    if vArea == 'Trn / Qld':
        texto(f'Área selecionada foi {vArea}. Clique em "Lançar demandas" para prosseguir.', 7)
        LinkArea = 'http://redminealgar/projects/gep_qldtrn/time_entries/new'
                  #'https://redmine.algartech.com/projects/gep_qldtrn/time_entries/new'
    if vArea == 'BI':
        texto(f'Área selecionada foi {vArea}. Clique em "Lançar demandas" para prosseguir.', 7)
        LinkArea = 'http://redminealgar/projects/mis-projetos-bi/time_entries/new'
                  #'https://redmine.algartech.com/projects/mis-projetos-bi/time_entries/new'
        
    b_lanca = Button(app, text='Lançar demandas', command = lambda: demandas(LinkArea))
    b_lanca.grid(row=9, column=1, pady=5, sticky=W)


# ### Valida Operações

# In[38]:


def valida_Opr(opr, lis):
    op = []
    for o in opr.split(','):
        #print(o)
        op.append(o)
    
    for operacao in op:
        if operacao not in lis:
            return('NOK')


# ### Lança demandas

# In[39]:


def demandas(LinkArea):
    
    try:
        #site = urllib.request.urlopen('http://redminealgar/projects/gep_qldtrn/time_entries/new')
        site= urllib.request.urlopen(LinkArea)
    except:
        texto('ERRO: O site do REDMINE está indisponivel, verifique a internet ou proxy!. Caso contrario contate o desenvolvedor (Michael Resende)!', 10)
    else:
        texto_branco(10)

#---------------------------------------------------------------------------------------------------------------#
#                                      Fazendo login                                                            #
#---------------------------------------------------------------------------------------------------------------#
        
        login = str(c_login.get())
        senha = str(c_senha.get())

#        chrome_opt = webdriver.ChromeOptions()
#        chrome_opt.add_argument('headless')
        #Abre o navegador
        navegador = webdriver.Chrome(executable_path=r'./chromedriver.exe')
#        navegador = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_opt)
        navegador.get(LinkArea)
        navegador.find_element_by_id('username').send_keys(login)
        navegador.find_element_by_id('password').send_keys(senha)
        navegador.find_element_by_name('login').send_keys(Keys.ENTER)
    
#---------------------------------------------------------------------------------------------------------------#
#                                      Acessando arquivo                                                        #
#---------------------------------------------------------------------------------------------------------------#

#tb_Demandas
        tb_Demanda = pd.read_excel('Demandas.xlsx')
        cont = tb_Demanda['Tarefa'].count()
    
#---------------------------------------------------------------------------------------------------------------#
#                                      Validando operações REDMINE                                              #
#---------------------------------------------------------------------------------------------------------------#

        lis = []
        op = []

#Lista as operações do Redmine
        for i in range(1, 301):
            try:
                elementos = navegador.find_element_by_xpath(f'//*[@id="time_entry_custom_field_values_18"]/option[{i}]').text
                lis.append(elementos)
                #print(elementos)
            except:
                break

#---------------------------------------------------------------------------------------------------------------#
#                                      Lançando demandas                                                        #
#---------------------------------------------------------------------------------------------------------------#

        for c in range(0, cont):
            if c == 0:
                validador = 0
                contador = 0
                try:
                    #Valida área
                    try:
                        navegador.find_element_by_id('time_entry_issue_id').clear()
                    except:
                        texto(f'ERRO: Sem autorização para lançamentos na Área selecionada! Avalie e tente novamente.', 10)
                        validador = 1
                        break
                    else:
                        #Valida Operação
                        if valida_Opr(tb_Demanda['Operação'].loc[0], lis) == 'NOK':
                            texto(f'ERRO: Operação invalida: {tb_Demanda["Operação"].loc[0]} - Data: {str(tb_Demanda["Data"].loc[0])[:10]}', 10)
                            validador = 1
                            break

                        #Tarefa
                        navegador.find_element_by_id('time_entry_issue_id').send_keys(int(tb_Demanda['Tarefa'].loc[0]))
                        #Data
                        navegador.find_element_by_id('time_entry_spent_on').clear()
                        navegador.find_element_by_id('time_entry_spent_on').send_keys(str(tb_Demanda['Data'].loc[0])[:10])
                        #Horas
                        navegador.find_element_by_id('time_entry_hours').clear()
                        navegador.find_element_by_id('time_entry_hours').send_keys(str(tb_Demanda['Horas'].loc[0])[:5])
                        #Comentario
                        navegador.find_element_by_id('time_entry_comments').clear()
                        navegador.find_element_by_id('time_entry_comments').send_keys(tb_Demanda['Comentario'].loc[0])
                        #Atividade
                        navegador.find_element_by_id('time_entry_activity_id').click()
                        navegador.find_element_by_id('time_entry_activity_id').send_keys(tb_Demanda['Atividade'].loc[0])

                        #Lista e clica nas operações da planilha
                        for o in tb_Demanda['Operação'].loc[0].split(','):
                            op.append(o)

                        for operacao in op:
                            for i, cr in enumerate(lis):
                                if operacao == cr:
                                    navegador.find_element_by_xpath(f'//*[@id="time_entry_custom_field_values_18"]/option[{i+1}]').click()

                        #Cliente
                        navegador.find_element_by_id('time_entry_custom_field_values_24').click()
                        navegador.find_element_by_id('time_entry_custom_field_values_24').send_keys(tb_Demanda['Cliente'].loc[0])
                        #Criar
                        navegador.find_element_by_name('continue').click()
                        sleep(2)
                        navegador.refresh()
                        op.clear()

                except:
                    texto(f'''ERRO: Tarefa invalida: {tb_Demanda["Comentario"].loc[0]}
                                Data: {str(tb_Demanda["Data"].loc[0])[:10]}''', 10)
                    validador = 1
                    break
                
                contador += 1
                tb_Demanda = tb_Demanda.drop(0)
                tb_Demanda.to_excel(f'Demandas.xlsx', index=False)
                
            else:
                tb_Demanda = pd.read_excel('Demandas.xlsx')
                try:
                    #Valida Operação
                    if valida_Opr(tb_Demanda['Operação'].loc[0], lis) == 'NOK':
                        texto(f'ERRO: Operação invalida: {tb_Demanda["Operação"].loc[0]} - Data: {str(tb_Demanda["Data"].loc[0])[:10]}', 10)
                        validador = 1
                        break

                    #Tarefa
                    navegador.find_element_by_id('time_entry_issue_id').clear()
                    navegador.find_element_by_id('time_entry_issue_id').send_keys(int(tb_Demanda['Tarefa'].loc[0]))
                    #Data
                    navegador.find_element_by_id('time_entry_spent_on').clear()
                    navegador.find_element_by_id('time_entry_spent_on').send_keys(str(tb_Demanda['Data'].loc[0])[:10])
                    #Horas
                    navegador.find_element_by_id('time_entry_hours').clear()
                    navegador.find_element_by_id('time_entry_hours').send_keys(str(tb_Demanda['Horas'].loc[0])[:5])
                    #Comentario
                    navegador.find_element_by_id('time_entry_comments').clear()
                    navegador.find_element_by_id('time_entry_comments').send_keys(tb_Demanda['Comentario'].loc[0])
                    #Atividade
                    navegador.find_element_by_id('time_entry_activity_id').click()
                    navegador.find_element_by_id('time_entry_activity_id').send_keys(tb_Demanda['Atividade'].loc[0])

                    #Lista e clica nas operações da planilha
                    for o in tb_Demanda['Operação'].loc[0].split(','):
                        op.append(o)

                    for operacao in op:
                        for i, cr in enumerate(lis):
                            if operacao == cr:
                                navegador.find_element_by_xpath(f'//*[@id="time_entry_custom_field_values_18"]/option[{i+1}]').click()

                    #Cliente
                    navegador.find_element_by_id('time_entry_custom_field_values_24').click()
                    navegador.find_element_by_id('time_entry_custom_field_values_24').send_keys(tb_Demanda['Cliente'].loc[0])
                    #Criar
                    navegador.find_element_by_name('continue').click()
                    sleep(2)
                    navegador.refresh()
                    op.clear()

                except:
                    texto(f'ERRO: Tarefa invalida: {tb_Demanda["Comentario"].loc[0]} - Data: {str(tb_Demanda["Data"].loc[0])[:10]}', 10)
                    validador = 1
                    break
                
                contador += 1
                tb_Demanda = tb_Demanda.drop(0)
                tb_Demanda.to_excel(f'Demandas.xlsx', index=False)

        navegador.quit()
        if validador == 1:
            texto(f'Foram lançados {contador} tarefas.',11)
        else:
            texto(f'Foram lançados {contador} tarefas.',10)
            texto_branco(11)
    


# ### Funções de texto

# In[40]:


def texto(string, l):
    tam = 0
    tam = 146 - len(string)
    if tam % 2 == 0:
        tam = int(tam / 2)
    else:
        tam += 1
        tam = int(tam / 2)
        
    aviso = f'{" "*(tam-1)} {string} {" "*(tam-1)}'
    aviso = Label(app, text=aviso.center(100))                        
    aviso.grid(row=l, column=0, columnspan=2)#, pady=5)
    
def texto_linha(l):
    branco = '-'*160
    branco = Label(app, text=branco.center(100))                        
    branco.grid(row=l, column=0, columnspan=2, pady=5)
    
def texto_branco(l):
    branco = ' '*280
    branco = Label(app, text=branco.center(100))                        
    branco.grid(row=l, column=0, columnspan=2, pady=5)


# ### Main

# In[41]:


app = Tk()
app.title('>>>> Lançador de demandas ---- Desenvolvido por: Michael Resende ---- michaeldmre@hotmail.com <<<<')
app.geometry('843x380')

t_login = Label(app, text='Login: ')
t_senha = Label(app, text='Senha: ')

c_login = Entry(app,)
c_senha = Entry(app, show='*')

b_login = Button(app, text='Logar', command = login)
#b_lanca = Button(app, text='Lançar demandas', command = demandas)
b_arq = Button(app, text='Buscar arquivo novamente', command = arquivo)


arquivo()

#app.resizable(False, False) #Não deixa redimencionar a janela

app.mainloop()


# In[ ]:




