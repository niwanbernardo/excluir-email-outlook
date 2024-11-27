import win32com.client
from win32com.client import Dispatch
import traceback

def excluir_emails(remetentes_para_excluir):
    try:
        # Conectando ao Outlook
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 representa a Caixa de Entrada
        
        # Recuperando os e-mails na Caixa de Entrada
        emails = inbox.Items  # Inclui e-mails lidos e não lidos
        print(f"Total de mensagens na caixa de entrada: {len(emails)}")

        # Inicializando contadores
        emails_excluidos = 0
        emails_verificados = 0

        # Iterar sobre as mensagens
        for email in emails:
            try:
                emails_verificados += 1
                remetente = email.SenderEmailAddress

                # Verifica se o remetente está na lista para exclusão
                if remetente in remetentes_para_excluir:
                    print(f"Excluindo e-mail de: {remetente}")
                    email.Delete()  # Excluindo o e-mail
                    emails_excluidos += 1

            except Exception as e:
                print(f"Erro ao processar um e-mail: {e}")
                continue

        print(f"\nProcessamento concluído!")
        print(f"Total de e-mails verificados: {emails_verificados}")
        print(f"Total de e-mails excluídos: {emails_excluidos}")

    except Exception as e:
        print("Erro ao acessar o Outlook ou processar e-mails.")
        print(traceback.format_exc())

# Lista de remetentes cujos e-mails serão excluídos
remetentes_para_excluir = [
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
    print("Iniciando exclusão de e-mails no Outlook...")
    excluir_emails(remetentes_para_excluir)
