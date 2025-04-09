import tkinter as tk
from tkinter import messagebox
import sys
from utils import abrir_arquivo, enviar_email

def exibir_tela_inicial(root, txt_feeback, cor_de_fundo_padrao):
    root.config(bg=cor_de_fundo_padrao,)
    txt_feeback.config(bg=cor_de_fundo_padrao)
    txt_feeback.config(text="Carregando...")


def exibir_feedback_planilha(root, planilhas=None, sucesso=False, txt_feeback=None, add_arquivo_btn=None):

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

     
def exibir_feedback_email(root, sucesso=False, txt_feeback=None, email_btn=None):

    if sucesso:
        root.config(bg="dark green")
        txt_feeback.config(bg="dark green", text="Emails enviados com sucesso✅")
        email_btn.pack_forget()
        
    else:
        root.config(bg="maroon")
        txt_feeback.config(bg="maroon")
        txt_feeback.config(text="Erro no envio dos emails❌\nNão foi possível realizar o envio\nFinalize e tente novamente")
        email_btn.pack_forget()   

    finalizar_btn = tk.Button(root, text=f"Finalizar", bg="ivory2", command=lambda: encerrar_script(root), font=("Arial", 16),  relief="raised")
    finalizar_btn.pack(side="bottom", padx=20, pady=40, expand=False, anchor="center")


def criar_btn_planilhas(root, objeto):
    # Botões para consultar as planilhas na tela
    btn1 = tk.Button(root, text=f"Abrir {objeto[0]["titulo"]}", bg="ivory2", command=lambda: abrir_arquivo(objeto[0]["diretorio"]), font=("Arial", 16),  relief="raised")
    btn1.pack(side="left", padx=20, pady=0, expand=True)

    btn2 = tk.Button(root, text=f"Abrir {objeto[1]["titulo"]}", bg="ivory2", command=lambda: abrir_arquivo(objeto[1]["diretorio"]), font=("Arial", 16),  relief="raised")
    btn2.pack(side="right", padx=20, pady=0, expand=True)


def criar_btn_email(root, txt_feeback):
    enviar_email_btn = tk.Button(root, text=f"Enviar emails", bg="ivory2", command=lambda: confirmar_envio_email(root, txt_feeback, enviar_email_btn), font=("Arial", 16),  relief="raised")
    enviar_email_btn.pack(side="bottom", padx=20, pady=40, expand=False, anchor="center")


def confirmar_envio_email(root, txt_feeback, email_btn):
    confirmacao = messagebox.askyesno("Confirmação de envio", "Enviar arquivos apresentados para otus e vanda?")

    if confirmacao:
        txt_feeback.config(text="Enviando emails...")

        try:
            enviar_email("Otus")
            enviar_email("Vanda")

            exibir_feedback_email(root=root, sucesso=True, txt_feeback=txt_feeback, email_btn=email_btn)

        except Exception as e:
            exibir_feedback_email(root=root, sucesso=False, txt_feeback=txt_feeback, email_btn=email_btn)


def encerrar_script(root):
    root.destroy() #Fecha o tkinter
    sys.exit(1)