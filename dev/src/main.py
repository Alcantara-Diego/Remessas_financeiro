import pandas as pd
import os
import sys
import tkinter as tk
from func_tkinter import exibir_feedback_usuario, exibir_tela_inicial
from alterar_planilha import tratarPlanilha, separar_otus_vanda
from controle_pastas import pedir_arquivo
from utils import escrever_log, enviar_email


def tkinter_config():
    root = tk.Tk()
    root.title("Separar Remessas")
    root.geometry("600x400")
    root.iconbitmap("pato.ico")
    cor_de_fundo_padrao = "skyblue4"
    root.config(bg=cor_de_fundo_padrao)

    return root, cor_de_fundo_padrao

root, cor_de_fundo_padrao = tkinter_config()


def encerrar_script():
    root.destroy() #Fecha o tkinter
    sys.exit(1)

def iniciar_script(cor_de_fundo_padrao):
    exibir_tela_inicial(root, txt_feeback, cor_de_fundo_padrao)

    caminho_arquivo = pedir_arquivo()

    try:  
        df_pagamentos = pd.read_excel(caminho_arquivo, sheet_name="Pagamentos")
        df_contas = pd.read_excel(caminho_arquivo, sheet_name="ContasBancarias")

        df_tratado = tratarPlanilha(df_pagamentos, df_contas)
        planilhas_separadas = separar_otus_vanda(df_tratado)

        exibir_feedback_usuario(root=root, planilhas=planilhas_separadas, sucesso=True, txt_feeback=txt_feeback, add_arquivo_btn=add_arquivo_btn)
        
    except Exception as e:
        exibir_feedback_usuario(root=root, sucesso=False, txt_feeback=txt_feeback, add_arquivo_btn=add_arquivo_btn)
        escrever_log("##########------------ERROR------------############", "a")
        escrever_log(e, "a")
    


 
# Criar texto de feecback
txt_feeback = tk.Label(root, text="Adicione um arquivo", bg=cor_de_fundo_padrao, fg="white", font=("Arial", 16, "bold"))
txt_feeback.pack(pady=20)  # "pack" coloca o widget na tela com um pouco de espaçamento

# Criar botão
add_arquivo_btn = tk.Button(root, text="Adicionar arquivo", bg="ivory2", command=lambda: iniciar_script(cor_de_fundo_padrao), font=("Arial", 16),  relief="raised")
add_arquivo_btn.pack(padx=20, pady=20, expand=True)

# Executa função após usuário fechar a tela
root.protocol("WM_DELETE_WINDOW", encerrar_script)

# Iniciar o loop da interface gráfica
root.mainloop()