import os
import platform
import glob
import base64
import pickle
import datetime
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


def abrir_arquivo(diretorio):

    if platform.system() == "Windows":
        os.startfile(diretorio)
    elif platform.system() == 'Darwin':  # macOS
        os.system(f'open {diretorio}')
    else:                                # Linux
        os.system(f'xdg-open {diretorio}')


def pegar_data_hoje():
    data_atual = datetime.date.today()
    data_formatada = data_atual.strftime("%d-%m-%Y")

    return data_formatada


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
    except Exception as e:
        escrever_log(f"Ocorreu um erro inesperado: {e}", "a")


def enviar_email(destinatario):
    from controle_pastas import pegar_pasta_atual #Evitar erro de importação circular
    escrever_log(f"---ENVIANDO EMAIL {destinatario}---", "a")

    service, email_usuario = autenticar_gmail()

    de = "me"
    para = "sistema1@guillaumon.com.br"
    assunto = "Teste de envio de email"
    corpo_email = f"Segue remessas {destinatario} para o período atual"
    data_hoje = pegar_data_hoje()

    try:
        pasta_atual = pegar_pasta_atual()
        pasta_remessas = f"{pasta_atual}/{destinatario}/*"

        arquivos = [os.path.abspath(f) for f in glob.glob(pasta_remessas) if (("otus" in os.path.basename(f).lower() or "vanda" in os.path.basename(f).lower()) and data_hoje in os.path.basename(f).lower()) and f.endswith(".xlsx")]
        escrever_log(arquivos, "a")

        mensagem_criada = criar_mensagem(de=de, para=para, assunto=assunto, corpo_email=corpo_email, caminho_arquivos=arquivos)

        if mensagem_criada:
            envio_realizado = realizar_envio(service=service, mensagem_criada=mensagem_criada)

        if envio_realizado:
            escrever_log(f"EMAIL ENVIADO COM SUCESSO\nDE {email_usuario} PARA {para}\nTITULO: {assunto}", "a")
            escrever_log(f"{envio_realizado}", "a")

    except Exception as e:
        escrever_log(f"Erro no envio do email: {e}", "a")
        raise e
