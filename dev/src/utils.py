import os
import platform
import datetime


def escrever_log(texto, tipo="a"):
    print(texto)
    with open('log.txt', tipo) as f:
        f.write(f"{texto}\n")


def abrir_arquivo(diretorio):

    if platform.system() == "Windows":
        os.startfile(diretorio)
    elif platform.system() == 'Darwin':  # mac
        os.system(f'open {diretorio}')
    else:                                # Linux
        os.system(f'xdg-open {diretorio}')


def pegar_data_hoje():
    data_atual = datetime.date.today()
    data_formatada = data_atual.strftime("%d-%m-%Y")

    return data_formatada

