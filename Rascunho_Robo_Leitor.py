import os
import tkinter as tk
from tkinter import ttk, messagebox
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
from email.utils import parsedate_to_datetime
# --- 1. NOVA FUNÇÃO DE AUTENTICAÇÃO (SÓ LEITURA) ---
def authenticate_gmail_readonly():
    """
    Autentica no Gmail com permissão ESTRITA APENAS PARA LEITURA.
    Salva as credenciais em 'token_leitura.json' para não conflitar com o token de envio.
    """
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    
    # Supondo que a função get_base_path() exista no código principal
    # Aqui usamos o diretório atual como base para o rascunho
    base_path = os.path.dirname(os.path.abspath(__file__)) 
    
    token_path = os.path.join(base_path, "token_leitura.json")
    secret_path = os.path.join(base_path, "AuthNOC.json")
    creds = None
    
    # O arquivo 'token_leitura.json' guarda os tokens de acesso e de atualização do usuário
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        
    # Se não houver credenciais (primeiro acesso) ou estiverem inválidas/expiradas
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(secret_path):
                raise FileNotFoundError(f"Arquivo '{secret_path}' não encontrado para autenticação.")
            
            # Realiza o login no navegador (pede permissão de leitura para o usuário)
            flow = InstalledAppFlow.from_client_secrets_file(secret_path, SCOPES)
            creds = flow.run_local_server(port=0)
            
        # Salva as credenciais para a próxima execução
        with open(token_path, 'w') as token: 
            token.write(creds.to_json())
            
    # Retorna o serviço do Gmail pronto para fazer buscas
    return build('gmail', 'v1', credentials=creds)


# --- 2. NOVA CLASSE DO APLICATIVO DO ROBÔ ---
class RoboLeitorGmailApp:
    def __init__(self, parent_frame, main_app_callback):
        self.frame = tk.Frame(parent_frame)
        self.frame.pack(fill="both", expand=True)
        self.main_app_callback = main_app_callback
        
        # O botão de voltar ao menu (Padrão da sua ferramenta)
        tk.Button(self.frame, text="VOLTAR AO MENU PRINCIPAL", command=self.main_app_callback, 
                  bg="#6c757d", fg="white", font=("Segoe UI", 9, "bold")).pack(pady=(5, 10))

        self.criar_interface()
        
        # Supondo que ThemeManager.apply_theme_to_all() exista no código principal
        # ThemeManager.apply_theme_to_all(self.frame)

    def criar_interface(self):
        # Container Principal
        self.container = tk.Frame(self.frame, padx=20, pady=20)
        self.container.pack(fill=tk.BOTH, expand=True)
        
        self.lbl_title = tk.Label(self.container, text="🤖 ROBÔ LEITOR DO GMAIL (Consulta de Respostas)", font=("Arial", 16, "bold"))
        self.lbl_title.pack(anchor="w", pady=(0, 20))
        
        # Área de Busca
        self.frame_busca = tk.Frame(self.container)
        self.frame_busca.pack(fill=tk.X, pady=(0, 15))
        
        self.lbl_busca = tk.Label(self.frame_busca, text="Termo de Busca (Ex: INEP 22047140):", font=("Arial", 11, "bold"))
        self.lbl_busca.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_buscar = tk.Button(self.frame_busca, text="🔍 BUSCAR E-MAILS", command=self.acao_buscar, 
                                    font=("Arial", 11, "bold"), bg="#0ea5e9", fg="white", cursor="hand2")
        self.btn_buscar.pack(side=tk.RIGHT, ipadx=10, ipady=4, padx=5)
        
        self.ent_busca = tk.Entry(self.frame_busca, font=("Arial", 12), width=40)
        self.ent_busca.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=4, padx=(0, 10))
        
        # Tabela de Resultados (Treeview)
        self.frame_tabela = tk.Frame(self.container)
        self.frame_tabela.pack(fill=tk.BOTH, expand=False, pady=(0, 15)) # Expand=False para dar espaço ao texto
        
        colunas = ("data", "remetente", "assunto")
        self.tree = ttk.Treeview(self.frame_tabela, columns=colunas, show="headings", height=5) # Menos linhas visíveis na tabela
        self.tree.heading("data", text="Data")
        self.tree.heading("remetente", text="Remetente")
        self.tree.heading("assunto", text="Assunto")
        
        self.tree.column("data", width=120, anchor="center")
        self.tree.column("remetente", width=250, anchor="w")
        self.tree.column("assunto", width=400, anchor="w")
        
        scrollbar_y = ttk.Scrollbar(self.frame_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Evento de clique na tabela
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # Área de Texto (Corpo do E-mail)
        self.lbl_corpo = tk.Label(self.container, text="Corpo da Mensagem Selecionada:", font=("Arial", 11, "bold"))
        self.lbl_corpo.pack(anchor="w", pady=(0, 5))
        
        self.txt_corpo = tk.Text(self.container, height=18, font=("Consolas", 10), wrap=tk.WORD)
        self.txt_corpo.pack(fill=tk.BOTH, expand=True)
        
        # Dicionário interno para guardar os corpos das mensagens baixadas (id_mensagem -> corpo_texto)
        self.mensagens_cache = {}

    def acao_buscar(self):
        termo = self.ent_busca.get().strip()
        if not termo:
            messagebox.showwarning("Aviso", "Por favor, digite um termo para buscar.")
            return
            
        # Limpar tabela e área de texto
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.txt_corpo.delete(1.0, tk.END)
        self.mensagens_cache.clear()
        
        self.btn_buscar.config(text="Buscando...", state="disabled")
        self.container.update()
        
        try:
            # 1. Autenticar no Gmail no modo leitura
            gmail_service = authenticate_gmail_readonly()
            
            # 2. Obter o perfil para saber qual é o seu e-mail
            perfil = gmail_service.users().getProfile(userId='me').execute()
            meu_email = perfil.get('emailAddress', '').lower()

            # 3. Realizar a busca na nuvem (API) limitando a 20 resultados
            resultados = gmail_service.users().messages().list(userId='me', q=termo, maxResults=20).execute()
            mensagens = resultados.get('messages', [])
            if not mensagens:
                msg_aviso = (
                    "Nenhum e-mail de solicitação foi localizado para este termo.\n\n"
                    "Para dar andamento ao pleito, é estritamente necessário o envio prévio do Mapa de Calor "
                    "(em formato PDF), o qual deve ser providenciado e disponibilizado pela equipe técnica "
                    "responsável pelo mapeamento da respectiva região."
                )
                messagebox.showinfo("Status da Solicitação", msg_aviso)
                return
                
            tem_resposta = False
                
            # 3. Baixar os detalhes de cada mensagem encontrada
            for msg in mensagens:
                dados_msg = gmail_service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
                payload = dados_msg.get('payload', {})
                headers = payload.get('headers', [])
                
                assunto = "(Sem Assunto)"
                remetente = "(Desconhecido)"
                data_recebimento = "(Data Desconhecida)"
                
                for h in headers:
                    if h['name'].lower() == 'subject':
                        assunto = h['value']
                    elif h['name'].lower() == 'from':
                        remetente = h['value']
                    elif h['name'].lower() == 'date':
                        data_recebimento = h['value']
                        
                # Formatar a data para dd/mm/aaaa
                try:
                    dt = parsedate_to_datetime(data_recebimento)
                    data_formatada = dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    data_formatada = data_recebimento[:20]
                        
                # Verificar se esse e-mail é uma resposta (remetente diferente do seu próprio)
                if meu_email not in remetente.lower():
                    tem_resposta = True

                # Extrair o corpo da mensagem (texto puro)
                corpo_texto = self.extrair_texto_da_mensagem(payload)
                
                # Montar o corpo com o cabeçalho 'De:' incluído
                corpo_exibicao = f"De: {remetente}\nData: {data_formatada}\n"
                corpo_exibicao += "-" * 50 + "\n\n"
                corpo_exibicao += corpo_texto
                
                self.mensagens_cache[msg['id']] = corpo_exibicao
                
                # Inserir na tabela (Treeview)
                self.tree.insert("", tk.END, iid=msg['id'], values=(data_formatada, remetente, assunto))
                
            # Finalizado o loop de leitura, verificar se achou alguma resposta de fora
            if not tem_resposta:
                msg_aguardando = (
                    "A solicitação de pleito foi devidamente enviada e encontra-se em análise.\n\n"
                    "No momento, aguardamos o parecer oficial por parte da equipe da EACE."
                )
                messagebox.showinfo("Status da Solicitação", msg_aguardando)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a busca:\n{str(e)}")
            
        finally:
            self.btn_buscar.config(text="🔍 BUSCAR E-MAILS", state="normal")

    def on_tree_select(self, event):
        selecionados = self.tree.selection()
        if not selecionados:
            return
            
        msg_id = selecionados[0]
        corpo = self.mensagens_cache.get(msg_id, "Corpo não disponível.")
        
        self.txt_corpo.delete(1.0, tk.END)
        self.txt_corpo.insert(tk.END, corpo)

    def extrair_texto_da_mensagem(self, payload):
        """
        Função recursiva para encontrar a parte de texto (text/plain) do payload do e-mail.
        O formato do Gmail é aninhado (multipart), então precisamos navegar pelas 'parts'.
        """
        try:
            mime_type = payload.get("mimeType")
            
            if mime_type == "text/plain":
                dados = payload.get("body", {}).get("data")
                if dados:
                    return base64.urlsafe_b64decode(dados).decode("utf-8", errors="ignore")
                    
            elif "parts" in payload:
                for part in payload["parts"]:
                    resultado = self.extrair_texto_da_mensagem(part)
                    if resultado:
                        return resultado
        except Exception as e:
            return f"[Erro ao extrair corpo da mensagem: {e}]"
            
        return "Nenhum texto puro encontrado na mensagem."

# (Abaixo dessa classe, na integração, eu colocaria um botão no Menu Principal apontando para RoboLeitorGmailApp)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Teste - Robô Leitor do Gmail")
    root.geometry("800x600")
    
    # Função dummy de callback para o botão "VOLTAR AO MENU PRINCIPAL"
    def dummy_callback():
        messagebox.showinfo("Voltar", "No programa principal, isso fecharia a tela e voltaria ao menu.")
        
    app = RoboLeitorGmailApp(root, dummy_callback)
    root.mainloop()
