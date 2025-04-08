import tkinter as tk
from tkinter import messagebox
import sys
from utils import abrir_arquivo, enviar_email

def exibir_tela_inicial(root, txt_feeback, cor_de_fundo_padrao):
    root.config(bg=cor_de_fundo_padrao,)
    txt_feeback.config(bg=cor_de_fundo_padrao)
    txt_feeback.config(text="Carregando...")


def exibir_feedback_usuario(root, planilhas=None, sucesso=False, txt_feeback=None, add_arquivo_btn=None):

    if planilhas is not None and sucesso:
        texto_feedback1 = f"Arquivos baixados✅\nPlanilha 1: {planilhas[0]["titulo"]}\nN° de linhas: {planilhas[0]["n_items"]}\n"
        texto_feedback2 = f"\nPlanilha 2: {planilhas[1]["titulo"]}\nN° de linhas: {planilhas[1]["n_items"]}\n"
 
        txt_feeback.config(text=texto_feedback1+texto_feedback2)

        add_arquivo_btn.pack(pady=0)
        add_arquivo_btn.pack_forget()

        criar_btn_planilhas(root, planilhas)
        criar_btn_email(root, txt_feeback)

    else:
        root.config(bg="maroon")
        add_arquivo_btn.pack()
        txt_feeback.config(bg="maroon")
        txt_feeback.config(text="Falha no tratamento da planilha❌\nArquivo enviado não está formatado corretamente\nTente novamente")
        

def criar_btn_planilhas(root, objeto):
    # Btn 1
    btn1 = tk.Button(root, text=f"Abrir {objeto[0]["titulo"]}", bg="ivory2", command=lambda: abrir_arquivo(objeto[0]["diretorio"]), font=("Arial", 16),  relief="raised")
    btn1.pack(side="left", padx=20, pady=0, expand=True)

    # Btn 2
    btn2 = tk.Button(root, text=f"Abrir {objeto[1]["titulo"]}", bg="ivory2", command=lambda: abrir_arquivo(objeto[1]["diretorio"]), font=("Arial", 16),  relief="raised")
    btn2.pack(side="right", padx=20, pady=0, expand=True)


def criar_btn_email(root, txt_feeback):
    enviar_email_btn = tk.Button(root, text=f"Enviar emails", bg="ivory2", command=lambda: confirmar_envio_email(root, txt_feeback, enviar_email_btn), font=("Arial", 16),  relief="raised")
    enviar_email_btn.pack(side="bottom", padx=20, pady=40, expand=False, anchor="center")


def confirmar_envio_email(root, txt_feeback, email_btn):
    confirmacao = messagebox.askyesno("Confirmação de envio", "Enviar arquivos apresentados para otus e vanda?")

    if confirmacao:
        enviar_email()
        root.config(bg="dark green")
        txt_feeback.config(bg="dark green", text="Emails enviados com sucesso✅")
        email_btn.pack_forget()

        encerrar_btn = tk.Button(root, text=f"Finalizar", bg="ivory2", command=lambda: encerrar_script(root), font=("Arial", 16),  relief="raised")
        encerrar_btn.pack(side="bottom", padx=20, pady=40, expand=False, anchor="center")






def encerrar_script(root):
    root.destroy() #Fecha o tkinter
    sys.exit(1)