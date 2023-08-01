import pandas as pd
import BD
import VerificaRepeticao
from tkinter import *
import smtplib
from email.message import EmailMessage
import usuarioEmail
import BD
import json
import verificarBD
from PIL import Image, ImageTk

####################     JSON       ####################
with open('base.json') as file:
    dados = json.load(file)

dadosBase = dados['base']

def EscreverJson(os):
   dadosBase.append({
    f"{len(dadosBase)}" : f"{os}"}) 
   gravar()
      
def gravar():
    with open("base.json", 'w') as f:
        json.dump(dados, f, indent= 1 )
  
def buscarInfoPGravar():
    for i in BD.bancoDeDados2:
        EscreverJson(i)

####################       EMAIL         ####################

#CONFIGURAR LOGIN E SENHA EMAIL
EMAIL_ADRESS = usuarioEmail.usuario
EMAIL_PASSWORD = usuarioEmail.senha

def ChamarAppEmail(): 
    for cliente in BD.bancoDeDados2:
        Msg(cliente)

#MONTANDO O EMAIL E ENVIANDO
def Msg(cliente):
    try:
        msg = EmailMessage()
        msg['From'] = f'{usuarioEmail.usuario}'
        msg['To'] = usuarioEmail.destinatario
        msg['Subject'] = f'{cliente} <assign sevengraficarapida@gmail.com>'
        msg.set_content('INSTALAÇÃO SEVEN')
        server = smtplib.SMTP('smtp-mail.outlook.com',587 )
        server.ehlo()
        server.starttls()
        server.login(EMAIL_ADRESS,EMAIL_PASSWORD)
        server.sendmail(msg['From'],msg['To'],msg.as_string())
        server.quit()
        TextOS['text'] = 'OK - Emails Enviados'
        buscarInfoPGravar()
    except:
        TextOS['text'] = 'Falha ao enviar emails! Provavel Quantidade de Emails Diarios ESGOTADOS!'

####################       EXCEL        ####################
# Lendo o arquivo Excel
df = pd.read_excel('AutoInstalacao.xlsx')

# FUNCAO VERIFICA SE É NUMERO
def isnumber(value,valor):
    try:
         float(value)
    except ValueError:
         return True
    return valor

# FUNÇAO PARA DIMINUIR CARACTERES
def diminuir(str,valor):
    max = 5
    if len(str) > max:
        return isnumber(str[:max],valor)
    else:
        return isnumber(str,valor)
# FUNÇAO PARA DIMINUIR CARACTERES DATA
def diminuirData(str):
    max = 10
    if len(str) > max:
        return str[:max] 
    else:
        return str
# FUNCAO INICIA A ANALISE DO EXCEL
def primeira():
    for i in range(1,len(df)):
        var = f"{df.iloc[i,0]}"
        var2 = f"{df.iloc[i,0]}"
        if diminuir(var,var2) == True:
            segunda(i,var2)
             
def segunda(i,var2):
    for b in range(i+1,len(df)):
        bot = f"{df.iloc[b,0]}"
        bot2 = f"{df.iloc[b,0]}"
        OS_Cliente = df.iloc[b,0]
        titulo = df.iloc[b,1]
        dataComHora = df.iloc[b,5]
        dataSemHora = diminuirData(f'{dataComHora}')
        if diminuir(bot,bot2) != True and bot != 'nan':
            final = str(df.iloc[i,0])+" "+str(OS_Cliente)+" "+str(titulo)+" "+f"<due {dataSemHora}>"
            BD.bancoDeDados.append(final)
            VerificaRepeticao.verificacao(final)
        elif diminuir(bot,bot2) == True: 
            return             
################## CHAMANDO AS PRIMEIRAS FUNCOES ###################

primeira()
verificarBD.verifica()

####################     INTERFACE GRAFICA       ####################
def listar():
    lb.delete(0,END)
    for index,i in enumerate(BD.bancoDeDados2):
        td = str(index)+" "+ i
        lb.insert(END, td)  

def chamaaa():
    ChamarAppEmail()
    
def deletes():
    BD.bancoDeDados2.pop(int(entrar.get()))
    lb.delete(0,END)
    listar()

#INTERFACE GRAFICA
janela = Tk()
janela.title('AutoTasks by LNeri')
p1 = PhotoImage(file = 'logo.png')
janela.iconphoto(False, p1)
textoPrincipal = Label(janela,text='Lista de OSs do Excel:',)
textoPrincipal.grid(column=0,row=0,padx=25,pady=0)
TextOS = Label(janela,text='',pady=25,wraplength=600,justify="left")
TextOS.grid(column=0,row=1,padx=25,pady=25)
botao = Button(janela,text='Buscar',command=listar)
botao.grid(column=0,row=2,padx=5,pady=5)
botao2 = Button(janela,text='Confirmar e Enviar',command=chamaaa)
botao2.grid(column=0,row=3,padx=5,pady=5)
lb = Listbox(janela,width=100,height=18,font=('Arial',11))
lb.grid(column=0,row=4,pady=10)
labelEntra = Label(janela,text='digite o indice do item que deseja apagar:')
labelEntra.grid(column=0,row=5,pady=10)
entrar = Entry(janela)
entrar.grid(column=0,row=6,pady=10)
botaoDelete = Button(janela,text='Excluir',command=deletes)
botaoDelete.grid(column=0,row=7,pady=10)
load = Image.open("logo.png")
render = ImageTk.PhotoImage(load)
img = Label(janela, image=render)
img.grid(column=0,row=7)
img.image = render
img.place(x=0, y=0)
janela.mainloop()