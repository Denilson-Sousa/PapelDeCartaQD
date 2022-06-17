# coding: utf-8

# # Projeto papel de carta outlook
# ## Montar o papel de carta com nome e cargo do colaborador
# 

# Pedir o nome do colaborador e cargo

import os
import win32com.client as win32
import re
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import askokcancel, showinfo, WARNING


def enviar_papel_carta(nome_informado, cargo_informado, email_informado):
    CaminhoCompleto = os.getcwd()

    #print("Entrada de dados")

    Nome = nome_informado
    Cargo = cargo_informado

    # Verifica se o endereço de e-mail é válido

    padrao_de_email = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    if re.fullmatch(padrao_de_email,email_informado):
        EmailColaborador = email_informado
        PosAsterisco = EmailColaborador.find("@")
        NomeUsuario = EmailColaborador[:PosAsterisco]
    else:
        #raise ValueError(f'Endereço de e-mail inválido {email_informado}')
        showinfo(title="Aviso", message='Endereço de e-mail inválido')
        return


    # Carrega o modelo de papel de carta e insere os dados do colaborador

    with open("Modelo.html", 'r') as ModeloBase:
        PapelUsuario = ModeloBase.read()

    PapelUsuario = PapelUsuario.replace("NomeSobrenome", Nome)
    PapelUsuario = PapelUsuario.replace("CargoFuncao", Cargo)

    # Gera novo papel de carta
    ArqNovoPapel = NomeUsuario + ".html"
    with open(ArqNovoPapel, 'w') as NovoPapel:
        NovoPapel.write(PapelUsuario)

    # Montando e enviando o email

    #print("Preparando e-mail")

    Outlook = win32.Dispatch("Outlook.Application")
    # Outlook = win32.gencache.EnsureDispatch("Outlook.Application")

    Mail = Outlook.CreateItem(0)
    Mail.To = EmailColaborador
    # Mail.bodyformat = 2
    Mail.Subject = "Configuração do Papel de Carta do Outlook"

    CorpoEmail = "Olá, " + Nome + """ <p><p>O arquivo anexado a esta mensagem é o seu papel de carta. Siga as instruções abaixo para configurar.</p></p>
    
    <p> 
    Para configurar o papel de carta do Outlook
    </p>
    
     <blockquote>• Abra o outlook, clique em "arquivo", "opções", "email", "papéis de carta e fontes", "Tema". Mas não selecione nenhum ainda. Apenas clique uma vez em "cancelar";</blockquote>
    
     <blockquote>• Pressione Win+R e digite %appdata%\Microsoft\Stationery e pressione ENTER. Salve o seu papel de carta nesta pasta;</blockquote>
    
     <blockquote>• Retorne ao Outlook e clique em "tema" novamente e selecione o papel de carta com seu nome. Clique "Ok" até retornar para a tela principal do Outlook;</blockquote>
    
    <br>
    <p>
    Para configurar o papel de carta em respostas/encaminhamentos
    </p>
    
    
      <blockquote>• Crie uma mensagem de e-mail e não digite nada. Apenas selecione todo o conteúdo e copie;</blockquote>
    
      <blockquote>• Em seguida, clique em "arquivo", "opções", "email", "assinaturas". Clique no botão "novo", dê um nome para a assinatura e clique ok. No painel inferior da janela, clique e pressione CTRL+V;</blockquote>
    
      <blockquote>• Depois, em "respostas/encaminhamentos" (e apenas aí), selecione a assinatura que criou. Clique "ok" até retornar à tela principal.</blockquote>
    
    <p><p>
    <b>Este email foi gerado automaticamente. Não responda.</b>
    </p></p>
    
    """

    Mail.HTMLBody = CorpoEmail

    NomeAnexo = CaminhoCompleto + "\\" + ArqNovoPapel
    #print(NomeAnexo)
    #print("Anexando...")
    Mail.Attachments.Add(NomeAnexo)

    # Mail.display(True)
    #print("Enviando...")
    Mail.Send()
    # Outlook.Quit()
    #print("Feito!")
    Outlook = None
    os.remove(NomeAnexo)

    showinfo(title='Aviso', message='Papel de carta enviado')
    exit()


# import tkinter as tk
# from tkinter import ttk
# from tkinter.messagebox import askokcancel, showinfo, WARNING

class Aplicativo(tk.Tk):
    def __init__(self):
        super().__init__()


        self.geometry('515x320')
        self.title('Emissão de papel de carta')

        self.resizable(0, 0)

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=8)

        self.create_widgets()

    def encerra_aplicativo(self):
        self.quit()

    def dispara_envio(self):
        enviar_papel_carta(self.nome.get(), self.cargo.get(),self.email.get())

    def create_widgets(self):
        self.nome = tk.StringVar()
        self.cargo = tk.StringVar()
        self.email = tk.StringVar()

        self.estilo = ttk.Style()
        self.estilo.theme_use('alt')
        self.estilo.configure('TButton', font=('Helvetica', 10))
        self.estilo.configure('TLabel', font=('Helvetica', 12))
        self.estilo.configure('Heading.TLabel', font=('Helvetica', 16))
        self.estilo.configure('TEntry', font=('Helvetica', 12))

        opcoes={'padx':15, 'pady':5, 'ipadx':10, 'ipady':10}

        titulo_pagina = ttk.Label(self, text="Informe os dados do colaborador",
                                  font=('Helvetica', 16),
                                  background='cyan',
                                  padding=(20,2),
                                  style='Heading.TLabel'
                                  )
        titulo_pagina.grid(column=0,
                           columnspan=2,
                           sticky=tk.N
                           )

        self.separador = ttk.Separator(self, orient='horizontal')
        self.separador.grid(column=0, row=1, columnspan=2)

        nome_label = ttk.Label(self,
                               text="Nome : ",
                               style='TLabel',
                               background="white")
        nome_label.grid(column=0,
                        row=4,
                        sticky=tk.W,
                        **opcoes)

        nome_entrada = ttk.Entry(self,
                                 textvariable=self.nome,
                                 width=100,
                                 background="blue",
                                 style='TEntry')
        nome_entrada.grid(column=1,
                          row=4,
                          sticky=tk.W,
                          **opcoes)

        cargo_label = ttk.Label(self,
                                text="Cargo: ",
                                style='TLabel',
                                background="white")
        cargo_label.grid(column=0,
                         row=5,
                         sticky=tk.W,
                         **opcoes)

        cargo_entrada = ttk.Entry(self,
                                  textvariable=self.cargo,
                                  width=100,
                                  style='TEntry')
        cargo_entrada.grid(column=1,
                           row=5,
                           sticky=tk.W,
                           **opcoes)

        email_label = ttk.Label(self,
                                text="e-mail : ",
                                style='TLabel',
                                background="white")
        email_label.grid(column=0,
                         row=6,
                         sticky=tk.W,
                         **opcoes)

        email_entrada = ttk.Entry(self,
                                  textvariable=self.email,
                                  width=100,
                                  style='TEntry',
                                  foreground='blue')
        email_entrada.grid(column=1,
                           row=6,
                           sticky=tk.W,
                           **opcoes)

        botao_enviar = ttk.Button(self,
                                  text='Enviar',
                                  width=10,
                                  style='TButton',
                                  command=lambda:self.dispara_envio())
        botao_enviar.grid(column=1,
                          row=7,
                          sticky=tk.E,
                          **opcoes)

        botao_fechar = ttk.Button(self,
                                  text='Fechar',
                                  width=10,
                                  command=lambda:self.encerra_aplicativo())
        botao_fechar.grid(column=1,
                          row=8,
                          sticky=tk.E,
                          **opcoes)

if __name__ == "__main__":
    aplicativo = Aplicativo()
    aplicativo.mainloop()



