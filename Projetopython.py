import os
import smtplib
import time
import pandas as pd
from tkinter import *
from tkinter import filedialog
from datetime import datetime
from email.message import EmailMessage
from filesecret import senhaEmail

data_atual = datetime.now()
dados_vencidos = None
data_formatada = data_atual.strftime("%d/%m/%Y")

Janela = Tk()
Janela.geometry('1000x500')
Janela.config(bg='#303030')
Janela.title('Sistema automatizado')

class Aplicacao:
    def carregar_relatorio(self):
        global dados_vencidos
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", "*")])
        if caminho:
            try:
                df = pd.read_excel(caminho)
                self.label_dados.delete(1.0, END)
                self.label_dados.insert(END, "CONTAS VENCIDAS\n\n")
                if 'VENCIMENTO' in df.columns and (pd.to_datetime(df['VENCIMENTO'], format="%d/%m/%Y") < data_atual).any():
                    dados_vencidos = df[pd.to_datetime(df['VENCIMENTO'], format="%d/%m/%Y") < data_atual]
                    self.label_dados.insert(END, dados_vencidos.to_string(index=False, header=False))
            except Exception as e:
                self.label_dados.delete(1.0, END)
                self.label_dados.insert(END, f"Erro ao ler arquivo: {str(e)}")  
    def Disparar_mensagens_Email(self):
        EMAIL_ADDRESS = 'paulocesarmartins2006@gmail.com'
        EMAIL_SENHA = senhaEmail
        
        for index,linha in dados_vencidos.iterrows():
            linhaProduto = linha['NOME DO PRODUTO']
            linhaValor = linha['PREÇO']
            linhaData = linha['VENCIMENTO'].strftime("%d/%m/%Y")
            linhaParcela = linha['N° DE PARCELAS']
            linhaCliente = linha['CLIENTE']
            linhaEmail = linha['EMAIL']
        
            msg = EmailMessage()
            msg['Subject'] = 'Cobrança'
            msg['From'] = 'paulocesarmartins2006@gmail.com'
            msg['To'] = linhaEmail

            conteudo_email = (
                f'Olá {linhaCliente},\n\n'
                f'Verificamos em nosso sistema uma parcela atrasada em seu nome. Por favor, entrar em contato para efetuar o pagamento.\n\n'
                f'Produto: {linhaProduto}\n'
                f'Valor: {linhaValor}\n'
                f'N° Parcela: {linhaParcela}\n'
                f'Vencimento: {linhaData}\n\n'
                f'Att\nSistema de Cobrança'
            )
            msg.set_content(conteudo_email)
            
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_SENHA)
                smtp.send_message(msg)
            time.sleep(1)
    def __init__(self, master=None):
        self.container = Frame(master)
        self.container.config(bg='#A7A7A7')
        self.container.pack()
        
        self.title = Label(text='SISTEMA DE COBRANÇA AUTOMATIZADO')
        self.title['font'] = 30
        self.title.config(bg='#303030',fg='white')
        self.title.pack()
        
        self.label_dados = Text(width=95, height=15)
        self.label_dados.config(bg='#DFDFDF',fg='black')
        self.label_dados.pack()
        
        self.bloco = Frame(master)
        self.bloco.config(bg='#303030')
        self.bloco.pack()
        
        self.start = Button(self.bloco,text='Carregar Relátorio',command=self.carregar_relatorio)
        self.start['width'] = 20
        self.start.config(bg='white',fg='black')
        self.start.pack(side='left',padx=5,pady=10)
        
        self.botao = Button(self.bloco,text='Disparar Cobrança',command=self.Disparar_mensagens_Email)
        self.botao.config(bg='white',fg='black')
        self.botao['width'] = 20
        self.botao.pack(side='left',padx=5)
        
Aplicacao(Janela)
Janela.mainloop()
