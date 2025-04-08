import tkinter as tk
import sys
from utils import abrir_arquivo


def exibir_tela_inicial(root, txt_feeback, cor_de_fundo_padrao):
    root.config(bg=cor_de_fundo_padrao,)
    txt_feeback.config(bg=cor_de_fundo_padrao)
    txt_feeback.config(text="Carregando...")


def exibir_feedback_usuario(root, planilhas=None, sucesso=False, txt_feeback=None, add_arquivo_btn=None):

    if planilhas is not None and sucesso:
        texto_feedback1 = f"Arquivos baixados✅\nPlanilha 1: {planilhas[0]["titulo"]}\nN° de linhas: {planilhas[0]["n_items"]}\n"
        texto_feedback2 = f"\nPlanilha 2: {planilhas[1]["titulo"]}\nN° de linhas: {planilhas[1]["n_items"]}\n"
 
        txt_feeback.config(text=texto_feedback1+texto_feedback2)

        add_arquivo_btn.pack_forget()

        criar_btn_abrir_planilha(root, planilhas[0])
        criar_btn_abrir_planilha(root, planilhas[1])

    else:
        root.config(bg="maroon")
        add_arquivo_btn.pack()
        txt_feeback.config(bg="maroon")
        txt_feeback.config(text="Falha no tratamento da planilha❌\nArquivo enviado não está formatado corretamente\nTente novamente")
        

def criar_btn_abrir_planilha(root, objeto):

    novo_btn = tk.Button(root, text=f"Abrir {objeto["titulo"]}", bg="ivory2", command=lambda: abrir_arquivo(objeto["diretorio"]), font=("Arial", 16),  relief="raised")
    novo_btn.pack(side="left", padx=20, pady=20, expand=True)


def encerrar_script(root):
    root.destroy() #Fecha o tkinter
    sys.exit(1)