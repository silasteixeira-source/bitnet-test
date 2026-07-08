import os
import json
from google_auth_oauthlib.flow import InstalledAppFlow

def gerar_token():
    print("="*50)
    print(" GERADOR DE TOKEN DE ACESSO - NOC EACE")
    print("="*50)
    print("\nEste programa irá abrir o navegador para que você faça login com sua conta do Google.")
    print("O token gerado permitirá que o enviador web mande e-mails no seu nome.\n")
    
    SCOPES_GMAIL = ['https://www.googleapis.com/auth/gmail.send']
    
    client_secret_file = "AuthNoc.json"
    
    if not os.path.exists(client_secret_file):
        print(f"[ERRO] O arquivo '{client_secret_file}' não foi encontrado na pasta atual.")
        print("Certifique-se de colocar este programa na mesma pasta que o AuthNoc.json.")
        input("\nPressione Enter para sair...")
        return
        
    try:
        flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES_GMAIL)
        creds = flow.run_local_server(port=0)
        
        token_filename = "meu_token_envio.json"
        with open(token_filename, 'w') as token:
            token.write(creds.to_json())
            
        print("\n" + "="*50)
        print(f"[SUCESSO] O arquivo '{token_filename}' foi gerado!")
        print("="*50)
        print("Faça o upload deste arquivo lá na página web (Streamlit) para habilitar o envio de e-mails em seu nome.")
    except Exception as e:
        print(f"\n[ERRO FATAL] Ocorreu um erro ao gerar o token:\n{e}")
        
    input("\nPressione Enter para sair...")

if __name__ == '__main__':
    gerar_token()
