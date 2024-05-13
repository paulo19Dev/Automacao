import smtplib
import time
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
from email.message import EmailMessage
from secretfile import emailpassword
from secretfile import email
from PIL import Image, ImageTk

data_atual = datetime.now()
dados_vencidos = None
data_formatada = data_atual.strftime("%d/%m/%Y")

Janela = Tk()
Janela.geometry('1360x900')
Janela.config(bg='#303030')
Janela.title('SISTEMA DE COBRANÇA AUTOMATIZADO')

class Aplicacao:
    def carregar_Relatorio(self):
        global dados_vencidos
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", "*")])
        if caminho:
            try:
                df = pd.read_excel(caminho)
                self.label_dados.delete(1.0, END)
                if 'VENCIMENTO' in df.columns and (pd.to_datetime(df['VENCIMENTO'], format="%d/%m/%Y") < data_atual).any():
                    dados_vencidos = df[pd.to_datetime(df['VENCIMENTO'], format="%d/%m/%Y") < data_atual]
                    self.label_dados.insert(END, dados_vencidos.to_string(index=False, header=False))
            except Exception as e:
                self.label_dados.delete(1.0, END)
                self.label_dados.insert(END, f"Erro ao ler arquivo: {str(e)}")
                
    def Disparar_Mensagem(self):
        EMAIL_ADDRESS = email
        EMAIL_SENHA = emailpassword
        
        for index,linha in dados_vencidos.iterrows():
            Produto = linha['NOME DO PRODUTO']
            Valor = linha['PREÇO']
            Data = linha['VENCIMENTO'].strftime("%d/%m/%Y")
            Parcela = linha['N° DE PARCELAS']
            Cliente = linha['CLIENTE']
            Email = linha['EMAIL']
        
            msg = EmailMessage()
            msg['Subject'] = 'Cobrança'
            msg['From'] = email
            msg['To'] = Email

            corpo_email = (
                f'Olá sr.{Cliente},\n\n'
                f'Verificamos em nosso sistema uma parcela atrasada em seu nome.\nDados da parcela:\n\n'
                f'Produto: {Produto}\n'
                f'Preço: R$ {Valor}\n'
                f'N° Parcela: {Parcela}\n'
                f'Vencimento: {Data}\n\n'
                f'Att\n SisCobra sistemas'
            )
            msg.set_content(corpo_email)
            
            messagebox.showinfo('SisCobra sistemas','Email enviado com sucesso!')
            
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_SENHA)
                smtp.send_message(msg)
            time.sleep(1)
            
    def limpar_Tela(self):
        self.label_dados.delete(1.0, END)
            
    def __init__(self, master=None):
    
        self.container = Frame(master)
        self.container.config(bg='#A7A7A7')
        self.container.pack()
        
        self.container0 = Frame(master)
        self.container0.config(bg='#303030')
        self.container0.pack()
        
        self.container1 = Frame(master)
        self.container1.config(bg='#303030')
        self.container1.pack()
        
        self.container2 = Frame(master)
        self.container2.config(bg='#303030')
        self.container2.pack()
        
        self.title = Label(self.container0,text='SISCOBRA SISTEMAS',font=("Arial", 20))
        self.title.config(bg='#303030',fg='white')
        self.title.pack(side=TOP,pady=30)
        
        self.label = Label(self.container1, text='CONTAS VENCIDAS', font=("Arial", 15))
        self.label.config(bg='#303030', fg='white')
        self.label.pack(padx=1, pady=10)

        self.label_dados = Text(self.container2, width=95, height=17)
        self.label_dados.config(bg='#DFDFDF', fg='black')
        self.label_dados.pack(side=LEFT, padx=1, pady=10)
        
        self.bloco = Frame(master)
        self.bloco.config(bg='#303030')
        self.bloco.pack()
        
        def logo():
            image = Image.open('img/logo.jpg')
            width, height = 100, 100
            image = image.resize((width, height))
            photo = ImageTk.PhotoImage(image)
            label = Label(Janela, image=photo)
            label.image = photo
            x_pos, y_pos = 850, 10
            label.place(x=x_pos, y=y_pos)  
        logo()
        
        self.start = Button(self.bloco,text='Abrir Relátorio',command=self.carregar_Relatorio)
        self.start['width'] = 20
        self.start.config(bg='white',fg='black')
        self.start.pack(side='left',padx=5,pady=10)
        
        self.botao = Button(self.bloco,text='Disparar Cobrança',command=self.Disparar_Mensagem)
        self.botao.config(bg='white',fg='black')
        self.botao['width'] = 20
        self.botao.pack(side='left',padx=5)
        
        self.limpar = Button(self.bloco,text='Limpar Tela',command=self.limpar_Tela)
        self.limpar.config(bg='white',fg='black')
        self.limpar['width'] = 20
        self.limpar.pack(side='left',padx=5)
        
Aplicacao(Janela)
Janela.mainloop()
