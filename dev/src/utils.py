import os
import base64
import pickle
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

# Defina o escopo para enviar e-mails
SCOPES = ['https://www.googleapis.com/auth/gmail.send','https://www.googleapis.com/auth/gmail.readonly']

def escrever_log(texto, tipo):
    print(texto)
    with open('log.txt', tipo) as f:
        f.write(f"{texto}\n")


def autenticar_gmail():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Salva o token para futuras execuções
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    profile = service.users().getProfile(userId='me').execute()
    email_usuario = profile.get('emailAddress')

    return service, email_usuario


def criar_mensagem(de, para, assunto, corpo_email, caminho_arquivo):
    mensagem = MIMEMultipart()
    mensagem['from'] = de
    mensagem['to'] = para
    mensagem['subject'] = assunto
    
    # Corpo do e-mail
    mensagem.attach(MIMEText(corpo_email, 'plain'))
    
    # Anexa o arquivo se ele existir
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
    except Exception as e:
        escrever_log(f"Ocorreu um erro inesperado: {e}", "a")


def enviar_email():
    service, email_usuario = autenticar_gmail()

    de = "me"
    para = "sistema1@guillaumon.com.br"
    assunto = "Teste de caso de uso"
    corpo_email = "Este é o corpo do e-mail."
    caminho_arquivo = '../../RelArquivoRemessa.xls'

    mensagem_criada = criar_mensagem(de=de, para=para, assunto=assunto, corpo_email=corpo_email, caminho_arquivo=caminho_arquivo)

    if mensagem_criada:
        envio_realizado = realizar_envio(service=service, mensagem_criada=mensagem_criada)

    if envio_realizado:
        escrever_log(f"EMAIL ENVIADO COM SUCESSO\nDE {email_usuario} PARA {para}\nTITULO: {assunto}", "a")
        escrever_log(f"{envio_realizado}", "a")


enviar_email()
