# Rascunho: Automação de Busca e Leitura de E-mails do Gmail

A ideia aqui é construir um robô autônomo focado em **Leitura de E-mails**. Assim como a sua ferramenta atual tem permissão para "Escrever e Enviar" (Gmail Send), nós podemos pedir permissão ao seu Google para "Apenas Leitura" (`gmail.readonly`). 

Isso vai nos dar um poder absurdo de puxar relatórios da sua própria caixa de entrada.

## 🛠️ Como Funciona o Motor de Busca (A Lógica)

O robô usará o método oficial do Google chamado `messages().list()`. Ele funciona mandando para o servidor do Gmail a exata mesma barra de pesquisa que um humano digitaria na web.

Por exemplo, nós podemos programar o robô para pesquisar:
* `q="subject:INEP 22047140"` (Procurar no título)
* `q="from:planejamento@eace.org.br 'aprovado'"` (Procurar texto de uma pessoa específica)
* `q="'Kit Wi-Fi' newer_than:7d"` (Procurar uma palavra-chave nos últimos 7 dias)

## 📦 O que o Programa Retorna?

Uma vez que o robô acha os e-mails, ele consegue abrir cada um (nos bastidores) e raspar as seguintes informações:
1. **O Remetente (De):** Quem mandou a mensagem.
2. **A Data e Hora:** Exatamente quando chegou.
3. **O Assunto:** Título da mensagem.
4. **O Texto do Corpo:** Nós podemos programar para ele ler o corpo do e-mail (HTML ou Texto Puro) e fatiar a resposta. Se o cliente respondeu "Aprovo a instalação", o robô captura essa frase e joga num arquivo texto, numa tela do aplicativo ou até preenche uma planilha do Sheets automaticamente!

## ⚙️ Exemplo Prático de Código do Rascunho

Se fossemos construir um esqueleto em Python puro hoje, ele seria mais ou menos assim:

```python
from googleapiclient.discovery import build

def buscar_respostas_email(palavra_chave):
    # 1. Autentica no Gmail com modo de Leitura
    gmail_service = build('gmail', 'v1', credentials=credenciais)
    
    # 2. Faz a busca usando a palavra chave
    # Exemplo de palavra chave: "INEP 22047140"
    query_busca = f"\"{palavra_chave}\" in:inbox"
    
    resultados = gmail_service.users().messages().list(
        userId='me', 
        q=query_busca
    ).execute()
    
    mensagens = resultados.get('messages', [])
    
    if not mensagens:
        print("Nenhuma resposta encontrada para essa escola/palavra.")
        return
        
    # 3. Varre as mensagens encontradas e puxa o conteúdo!
    for msg in mensagens:
        dados_msg = gmail_service.users().messages().get(
            userId='me', 
            id=msg['id'], 
            format='full'
        ).execute()
        
        # Extrair assunto e remetente
        cabecalhos = dados_msg['payload']['headers']
        assunto = next(h['value'] for h in cabecalhos if h['name'] == 'Subject')
        remetente = next(h['value'] for h in cabecalhos if h['name'] == 'From')
        
        print(f"📩 E-MAIL ENCONTRADO!")
        print(f"De: {remetente}")
        print(f"Assunto: {assunto}")
        
        # (O robô aqui pegaria o texto puro e salvaria pra você)
```

## Como Integrar no Novo Projeto
1. Peça para a IA incluir o escopo `https://www.googleapis.com/auth/gmail.readonly` no processo de login OAuth2.
2. Defina os inputs (Onde a pessoa vai digitar a palavra chave).
3. Defina a rota de saída (Onde o texto do email lido será jogado: no terminal, na UI, em uma tabela excel).
