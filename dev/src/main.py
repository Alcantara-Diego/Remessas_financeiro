import pandas as pd
import sys
from utils import escrever_log, enviar_email
from alterar_planilha import tratarPlanilha, separar_otus_vanda
from controle_pastas import pedir_arquivo


caminho_arquivo = pedir_arquivo()


try:
    df_pagamentos = pd.read_excel(caminho_arquivo, sheet_name="Pagamentos")
    df_contas = pd.read_excel(caminho_arquivo, sheet_name="ContasBancarias")

    df_tratado = tratarPlanilha(df_pagamentos, df_contas)

    separar_otus_vanda(df_tratado)
    
except ValueError as e:
    escrever_log("##########------------ERROR------------############", "a")
    escrever_log(e, "a")
    sys.exit(1)