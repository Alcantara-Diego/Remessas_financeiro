import pandas as pd
import os
from utils import escrever_log
from controle_pastas import criar_pasta_remessas

def tratarPlanilha(df_pagamentos, df_contas):
    print("---")

    df_mesclado = mesclarPlanilhas(df_pagamentos, df_contas)

    # Apagar colunas que o setor nn vai precisar
    df_tratado = df_mesclado.drop(['codigo_logo', 'sistema', 'id_conta_bancaria', 'bloco', 'vencto_ok', 'valor_ok', 'tem_docto', 'link', 'catalogos_id'], axis=1)

    df_formatado = formatar_data(df_tratado)

    return df_formatado


def separar_otus_vanda(df_base):
    print("------")

    print("SEPARANDO ARQUIVOS....")
    df_otus = df_base[df_base["codigo"] <= 356]
    df_vanda = df_base[df_base["codigo"] >= 357]

    recebedores = [
        {"titulo": "Otus", "planilha": df_otus, "n_items": 0, "diretorio": ""},
        {"titulo": "Vanda", "planilha": df_vanda, "n_items": 0, "diretorio": ""}
    ]

    caminho_remessas = criar_pasta_remessas()

    for recebedor in recebedores:
        titulo = recebedor["titulo"]
        planilha = recebedor["planilha"]
        
        novo_arquivo = os.path.join(caminho_remessas, f"{titulo}({len(planilha)})_RelArquivoRemessa.xlsx")
        recebedor["n_items"] = len(planilha)
        recebedor["diretorio"] = os.path.abspath(novo_arquivo)
        
        planilha.to_excel(novo_arquivo, index=False)

        print(f"ARQUIVO - {titulo} criado com sucesso!")
        escrever_log(f'Caminho do arquivo criado: {recebedor["diretorio"]}\n', "a")

    return recebedores

def mesclarPlanilhas(df_pagamentos, df_contas):
    print("mergin....")

    df_pagamentos.insert(4, 'descricao', "")
  
    df_combinado = df_pagamentos.merge(df_contas[["descricao", "codigo"]], left_on="conta_bancaria", right_on="codigo", how="left")
    df_pagamentos["descricao"] = df_combinado["descricao_y"]

    return df_pagamentos
    
def formatar_data(df):
    print("Formatando data.....")

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
