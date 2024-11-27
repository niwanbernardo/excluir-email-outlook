import win32com.client
from win32com.client import Dispatch
import traceback

def recuperar_emails(remetentes_para_recuperar):
    try:
        # Conectando ao Outlook
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Acessando a pasta de Itens Excluídos (3 representa "Deleted Items")
        itens_excluidos = outlook.GetDefaultFolder(3)
        
        # Acessando a Caixa de Entrada (6 representa "Inbox")
        caixa_de_entrada = outlook.GetDefaultFolder(6)
        
        # Recuperando os e-mails da pasta de Itens Excluídos
        emails = itens_excluidos.Items
        print(f"Total de mensagens nos Itens Excluídos: {len(emails)}")
        
        # Inicializando contadores
        emails_recuperados = 0
        emails_verificados = 0

        # Iterar sobre as mensagens
        for email in emails:
            try:
                emails_verificados += 1
                remetente = email.SenderEmailAddress

                # Verifica se o remetente está na lista para recuperação
                if remetente in remetentes_para_recuperar:
                    print(f"Recuperando e-mail de: {remetente}")
                    email.Move(caixa_de_entrada)  # Movendo o e-mail de volta para a Caixa de Entrada
                    emails_recuperados += 1

            except Exception as e:
                print(f"Erro ao processar um e-mail: {e}")
                continue

        print(f"\nProcessamento concluído!")
        print(f"Total de e-mails verificados: {emails_verificados}")
        print(f"Total de e-mails recuperados: {emails_recuperados}")

    except Exception as e:
        print("Erro ao acessar o Outlook ou processar e-mails.")
        print(traceback.format_exc())

# Lista de remetentes cujos e-mails serão recuperados
remetentes_para_recuperar = [
    "spam1@example.com",
    "spam2@example.com",
    "newsletter@spam.com",
    # E-mails do LinkedIn (comentados para fácil inclusão, se necessário)
    # "invitations@linkedin.com",
    # "notifications-noreply@linkedin.com",
    # "jobs-listings@linkedin.com",
    # "jobalerts-noreply@linkedin.com",
    # "messages-noreply@linkedin.com",
    # "linkedin@e.linkedin.com",
    # "editors-noreply@linkedin.com",
]

# Executando a função
if __name__ == "__main__":
    print("Iniciando recuperação de e-mails nos Itens Excluídos...")
    recuperar_emails(remetentes_para_recuperar)
