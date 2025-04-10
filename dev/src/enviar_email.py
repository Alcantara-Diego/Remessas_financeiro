import os
import sys
import glob
import base64
import pickle
from utils import escrever_log
from controle_pastas import pegar_pasta_atual, renomear_arquivos
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# Defina o escopo para enviar e-mails
SCOPES = ['https://www.googleapis.com/auth/gmail.send','https://www.googleapis.com/auth/gmail.readonly']


def autenticar_gmail():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            
            if getattr(sys, 'frozen', False):
                # Se for empacotado, o arquivo 'credentials.json' estará no diretório temporário
                caminho_credenciais = os.path.join(sys._MEIPASS, 'credentials.json')
            else:
                # Se for executado diretamente, o arquivo 'credentials.json' estará no diretório do script
                caminho_credenciais = 'credentials.json'

            flow = InstalledAppFlow.from_client_secrets_file(caminho_credenciais, SCOPES)
            creds = flow.run_local_server(port=0)

        # Salva o token para futuras execuções
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    profile = service.users().getProfile(userId='me').execute()
    email_usuario = profile.get('emailAddress')

    return service, email_usuario


def criar_mensagem(de, para, assunto, corpo_email, caminho_arquivos):
    mensagem = MIMEMultipart()
    mensagem['from'] = de
    mensagem['to'] = para
    mensagem['subject'] = assunto
    
    # Corpo do e-mail
    mensagem.attach(MIMEText(corpo_email, 'plain'))
    
    # Anexa o arquivo se ele existir
    for caminho_arquivo in caminho_arquivos:
        if os.path.exists(caminho_arquivo):
            with open(caminho_arquivo, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(caminho_arquivo))
                mensagem.attach(part)
    else:
        escrever_log(f"Não foi possível encontrar o arquivo em: {caminho_arquivo}", "a")
    
    return {'raw': base64.urlsafe_b64encode(mensagem.as_bytes()).decode()}


def realizar_envio(service, mensagem_criada):
    try:
        mensagem_enviada = service.users().messages().send(userId="me", body=mensagem_criada).execute()
        return mensagem_enviada
    except HttpError as error:
        escrever_log(f"Ocorreu um erro ao tentar enviar o e-mail: {error}", "a")
        raise error
    except Exception as e:
        escrever_log(f"Ocorreu um erro inesperado: {e}", "a")
        raise e


def enviar_email(destinatario):
    escrever_log(f"---ENVIANDO EMAIL {destinatario}---", "a")

    de = "me"
    para = "sistema1@guillaumon.com.br"
    assunto = "Teste de envio de email"
    corpo_email = f"Segue remessas {destinatario} para o período atual"

    try:
        service, email_usuario = autenticar_gmail()

        pasta_atual = pegar_pasta_atual()
        pasta_remessas = f"{pasta_atual}/{destinatario}/*"

        arquivos = [os.path.abspath(f) for f in glob.glob(pasta_remessas) if (("otus" in os.path.basename(f).lower() or "vanda" in os.path.basename(f).lower()) and "enviado" not in os.path.basename(f).lower()) and f.endswith(".xlsx")]
        escrever_log(arquivos, "a")

        mensagem_criada = criar_mensagem(de=de, para=para, assunto=assunto, corpo_email=corpo_email, caminho_arquivos=arquivos)

        if mensagem_criada:
            envio_realizado = realizar_envio(service=service, mensagem_criada=mensagem_criada)

        if envio_realizado:
            renomear_arquivos(arquivos) #Marca-los como enviados

            escrever_log(f"EMAIL ENVIADO COM SUCESSO\nDE {email_usuario} PARA {para}\nTITULO: {assunto}", "a")
            escrever_log(f"{envio_realizado}", "a")

    except Exception as e:
        escrever_log(f"Erro no envio do email: {e}", "a")
        raise e
