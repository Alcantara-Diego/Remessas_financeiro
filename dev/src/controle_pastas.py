import os
import sys
import argparse
import traceback
from utils import pegar_data_hoje
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from utils import escrever_log

def selecionar_arquivo():
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    caminho_arquivo = askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xls;*.xlsx")],
    )

    escrever_log(f'Caminho do arquivo recebido: {caminho_arquivo}', "a")

    return caminho_arquivo


def pedir_arquivo():
    escrever_log("Iniciando script....", "w")

    parser = argparse.ArgumentParser(description="Processar arquivo Excel de remessa.")
    parser.add_argument("arquivo", nargs="?", help="Caminho para o arquivo Excel a ser processado.")
    args = parser.parse_args()

    if not args.arquivo:
        caminho_arquivo = selecionar_arquivo()
        
        if not caminho_arquivo:
            print("Nenhum arquivo foi selecionado")

    else:
        caminho_arquivo = args.arquivo

    if not os.path.exists(caminho_arquivo):
        print(f"Erro: O arquivo {caminho_arquivo} não foi encontrado.")

    print("Caminho do arquivo: ", caminho_arquivo)
    return caminho_arquivo


def criar_pasta(titulo):

    pasta_atual = pegar_pasta_atual()

    caminho_diretorio = os.path.join(pasta_atual, titulo)

    if not os.path.exists(caminho_diretorio):
        os.makedirs(caminho_diretorio)

    escrever_log(f'Pasta {titulo} criada em: {caminho_diretorio}', "a")

    return caminho_diretorio


def pegar_pasta_atual():

    if getattr(sys, 'frozen', False):
        caminho_script = os.getcwd() #caminho para arquivo .exe
    else:
        caminho_script = os.path.abspath(os.path.dirname(__file__)) #caminho para .py no vscode

    return caminho_script

    
def renomear_arquivos(lista_arquivos):
    data_hoje = pegar_data_hoje()

    for arquivo in lista_arquivos:

        try:
            novo_nome = arquivo.replace(".xlsx", f"_enviado_{data_hoje}.xlsx")

            #Evitar erro com nomes duplicados
            contador = 2
            while os.path.exists(novo_nome):
                novo_nome = arquivo.replace(".xlsx", f"_enviado_{data_hoje}_v{contador}.xlsx")
                contador += 1

            os.rename(arquivo, novo_nome)
            escrever_log(f"Arquivo renomeado: {novo_nome}")

        except Exception as e:
            detalhes_erro = traceback.format_exc()
            escrever_log(f"### ERRO AO RENOMEAR ARQUIVO: {detalhes_erro} ###", "a")




