import pandas as pd
import tkinter as tk
from func_tkinter import exibir_feedback_planilha, exibir_tela_inicial, encerrar_script
from alterar_planilha import mesclarPlanilhas, tratarPlanilha, separar_otus_vanda, baixar_planilhas
from controle_pastas import pedir_arquivo
from utils import escrever_log


def tkinter_config():
    root = tk.Tk()
    root.title("Separar Remessas")
    root.geometry("600x400")
    cor_de_fundo_padrao = "skyblue4"
    root.config(bg=cor_de_fundo_padrao)
    root.protocol("WM_DELETE_WINDOW", lambda: encerrar_script(root))

    return root, cor_de_fundo_padrao


root, cor_de_fundo_padrao = tkinter_config()


def iniciar_script(cor_de_fundo_padrao):
    exibir_tela_inicial(root, txt_feeback, cor_de_fundo_padrao)

    caminho_arquivo = pedir_arquivo()

    try:  
        df_pagamentos = pd.read_excel(caminho_arquivo, sheet_name="Pagamentos")
        df_contas = pd.read_excel(caminho_arquivo, sheet_name="ContasBancarias")
        
        # Pega os dados importantes do df_contas para inserir no df_pagamentos
        df_mesclado = mesclarPlanilhas(df_pagamentos, df_contas)
        # Apaga colunas não necessárias e formata a data corretamente
        df_tratado = tratarPlanilha(df_mesclado)

        planilhas_separadas = separar_otus_vanda(df_tratado)

        planilhas_baixadas = baixar_planilhas(planilhas_separadas)

        exibir_feedback_planilha(root=root, planilhas=planilhas_baixadas, sucesso=True, txt_feeback=txt_feeback, add_arquivo_btn=add_arquivo_btn)
        
    except Exception as e:
        exibir_feedback_planilha(root=root, sucesso=False, txt_feeback=txt_feeback, add_arquivo_btn=add_arquivo_btn)
        escrever_log("##########------------ERROR------------############", "a")
        escrever_log(e, "a")
    


 
# Criar texto de feecback
txt_feeback = tk.Label(root, text="Adicione um arquivo", bg=cor_de_fundo_padrao, fg="white", font=("Arial", 16, "bold"))
txt_feeback.pack(pady=20)  # "pack" coloca o widget na tela com um pouco de espaçamento

# Criar botão pedir arquivo
add_arquivo_btn = tk.Button(root, text="Adicionar arquivo", bg="ivory2", command=lambda: iniciar_script(cor_de_fundo_padrao), font=("Arial", 16),  relief="raised")
add_arquivo_btn.pack(padx=20, pady=0, expand=True)

# Inicia o tkinter
root.mainloop()