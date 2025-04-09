import pandas as pd
import os
import datetime
from utils import escrever_log, pegar_data_hoje
from controle_pastas import criar_pasta


def mesclarPlanilhas(df_pagamentos, df_contas):
    escrever_log("---JUNTANDO PLANILHAS---", "a")

    df_pagamentos.insert(4, 'descricao', "")
  
    df_combinado = df_pagamentos.merge(df_contas[["descricao", "codigo"]], left_on="conta_bancaria", right_on="codigo", how="left")
    df_pagamentos["descricao"] = df_combinado["descricao_y"]

    return df_pagamentos
    

def tratarPlanilha(df):
    escrever_log("---TRATANDO PLANILHAS---", "a")

    # Apagar colunas que o setor nn vai precisar
    df_tratado = df.drop(['codigo_logo', 'sistema', 'id_conta_bancaria', 'bloco', 'vencto_ok', 'valor_ok', 'tem_docto', 'link', 'catalogos_id'], axis=1)

    df_formatado = formatar_data(df_tratado)

    return df_formatado


def formatar_data(df):
    escrever_log("---FORMATANDO DATAS---", "a")

     # Verifica se 'vencimento' está no tipo datetime, se não, converte
    if not pd.api.types.is_datetime64_any_dtype(df['vencimento']):
        df['vencimento'] = pd.to_datetime(df['vencimento'], errors='coerce', origin='1899-12-30', unit='D')
    # 1462 ajusta a data de vencimento para a data correta
    df['vencimento'] = df['vencimento'] + pd.to_timedelta(1462, unit='D')
    df['vencimento'] = df['vencimento'].dt.strftime('%d/%m/%Y')


    if not pd.api.types.is_datetime64_any_dtype(df['data_pagto']):
        df['data_pagto'] = pd.to_datetime(df['data_pagto'], errors='coerce', origin='1899-12-30', unit='D')

    df['data_pagto'] = df['data_pagto'] + pd.to_timedelta(1462, unit='D')
    df['data_pagto'] = df['data_pagto'].dt.strftime('%d/%m/%Y')

    return df


def separar_otus_vanda(df_base):
    escrever_log("---SEPARANDO ARQUIVOS---", "a")

    df_otus = df_base[df_base["codigo"] <= 356]
    df_vanda = df_base[df_base["codigo"] >= 357]

    array_destinatarios = [
        {"titulo": "Otus", "planilha": df_otus, "n_items": 0, "diretorio": ""},
        {"titulo": "Vanda", "planilha": df_vanda, "n_items": 0, "diretorio": ""}
    ]

    return array_destinatarios


def baixar_planilhas(array_destinatarios):
    escrever_log("---BAIXANDO ARQUIVOS---", "a")    
    
    for destinatario in array_destinatarios:
        titulo = destinatario["titulo"]
        planilha = destinatario["planilha"]
        
        diretorio = criar_pasta(titulo)
        data_hoje = pegar_data_hoje()

        novo_arquivo = os.path.join(diretorio, f"{titulo}({len(planilha)})_{data_hoje}.xlsx")
        destinatario["n_items"] = len(planilha)
        destinatario["diretorio"] = os.path.abspath(novo_arquivo)
        
        planilha.to_excel(novo_arquivo, index=False)

        escrever_log(f"ARQUIVO - {titulo} criado com sucesso!\n", "a")
        escrever_log(f'Caminho do arquivo criado: {destinatario["diretorio"]}\n', "a")

    return array_destinatarios

