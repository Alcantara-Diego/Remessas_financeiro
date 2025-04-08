import os
import sys
import argparse
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


def criar_pasta_remessas():

    if getattr(sys, 'frozen', False):
        caminho_script = os.getcwd()
    else:
        caminho_script = os.path.abspath(os.path.dirname(__file__))

    caminho_remessas = os.path.join(caminho_script, 'RemessasOtusVanda')

    if not os.path.exists(caminho_remessas):
        os.makedirs(caminho_remessas)

    escrever_log(f'Pasta remessas criada em: {caminho_remessas}', "a")
    escrever_log(f'Script sendo aberto em: {caminho_script}', "a")
    

    return caminho_remessas


