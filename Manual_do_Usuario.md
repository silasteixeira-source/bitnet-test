# Manual do Usuário — Ferramenta EACE NOC (v6.1)
*Bitnet Telecom — Guia Completo de Instalação, Configuração e Operação*

---

## 📑 Índice
1. [Visão Geral do Sistema](#1-visão-geral-do-sistema)
2. [Guia de Instalação e Atualização](#2-guia-de-instalação-e-atualização)
3. [Pré-requisitos e Arquivos de Configuração](#3-pré-requisitos-e-arquivos-de-configuração)
4. [Guia Passo a Passo das Ferramentas](#4-guia-passo-a-passo-das-ferramentas)
   - [4.1 Gerador de Contrato EACE](#41-gerador-de-contrato-eace)
   - [4.2 Consulta EACE & Verificador de Velocidade](#42-consulta-eace--verificador-de-velocidade)
   - [4.3 Unificador de Evidências Multiformato](#43-unificador-de-evidências-multiformato)
   - [4.4 Conversor Bidirecional PDF ⇄ Word](#44-conversor-bidirecional-pdf--word)
   - [4.5 Dashboard de Monitoramento EACE](#45-dashboard-de-monitoramento-eace)
5. [Solução de Problemas (Troubleshooting)](#5-solução-de-problemas-troubleshooting)

---

## 1. Visão Geral do Sistema
A **Ferramenta EACE NOC** é um ecossistema de produtividade desenvolvido em Python (Tkinter) para automatizar e otimizar as atividades do NOC da Bitnet Telecom. O sistema integra inteligência local, cache resiliente e comunicação em tempo real com planilhas hospedadas no Google Sheets. Ele está projetado com um visual moderno (temas claro e escuro) e micro-animações para garantir uma excelente experiência de usuário.

---

## 2. Guia de Instalação e Atualização

### 📥 Como Instalar o Aplicativo
O software possui um instalador simplificado para Windows que não exige privilégios de administrador.

1. Localize o arquivo de instalação **`Ferramenta_EACE_NOC_Setup_v6.1.exe`** (geralmente gerado na pasta `C:\InstaladorNOC\`).
2. Dê um duplo clique no instalador.
3. Se o Windows exibir um alerta de segurança (SmartScreen), clique em **"Mais informações"** e depois em **"Executar assim mesmo"**.
4. Avance pelas etapas do assistente:
   - O programa será instalado por padrão no diretório de usuário:  
     `C:\Users\<Seu_Usuario>\AppData\Local\Programs\FeramentaEACEBitnet`
   - Selecione a caixa **"Criar atalho na Área de Trabalho"** se desejar acesso rápido.
5. Clique em **Instalar** e, ao término, em **Concluir**. O aplicativo pode ser configurado para iniciar automaticamente.

### 🔄 Como Atualizar
Se uma nova versão for disponibilizada:
1. Feche o programa ativo.
2. Execute o instalador da nova versão (`v6.1` ou superior).
3. O instalador reconhece a versão anterior, substitui o executável principal e os modelos de contrato automaticamente, **preservando as suas credenciais (`credentials.json`) e caches locais** para que você não perca suas configurações de trabalho.

### 🗑️ Como Desinstalar
Caso precise remover o programa:
1. Abra o Menu Iniciar do Windows.
2. Procure pelo grupo **"Ferramenta EACE NOC"**.
3. Clique em **"Desinstalar Ferramenta EACE NOC"** (ou utilize o Painel de Controle > Adicionar ou Remover Programas).
4. As configurações salvas e os diretórios vazios serão removidos do computador de forma limpa.

---

## 3. Pré-requisitos e Arquivos de Configuração

Para usufruir de 100% dos recursos do software, a pasta de instalação necessita de alguns arquivos auxiliares, a maioria distribuída junto ao instalador:

| Arquivo | Descrição | Onde fica? | Observação |
| :--- | :--- | :--- | :--- |
| **`EACENOC.xlsx`** | Planilha permanente de cadastro de escolas. | Pasta do App | **Planilha oficial do projeto EACE** integrada ao instalador para uso permanente e automático no Gerador de Contrato. |
| **`credentials.json`** | Chave da Conta de Serviço do Google Cloud. | Pasta do App | Essencial para sincronizar os dados da nuvem em tempo real. Se ausente, o app operará em modo offline. |
| **`eace_cache.json`** | Cache local das escolas. | Pasta do App | Utilizado pelo módulo "Consulta EACE" caso não haja conexão com a internet. |
| **`eace_dashboard_cache.json`** | Cache local do monitoramento. | Pasta do App | Utilizado pelo módulo "Dashboard EACE". |
| **`contrato modeloBIT.docx`** | Modelo Word da contratada Bitnet. | Pasta do App | Modelo editável com marcações `{{tag}}` utilizado na geração de contratos. |
| **`contrato modeloST1.docx`** | Modelo Word da contratada ST1. | Pasta do App | Modelo editável complementar para geração de contratos. |
| **`config_apps.json`** | Salva preferências locais do usuário. | Pasta do App | Gerado automaticamente (guarda tema escolhido e pastas padrão de exportação). |

> [!IMPORTANT]
> **Microsoft Word:** Para usar a conversão de documentos Word (`.docx`) para PDF ou incluir arquivos Word no Unificador de Evidências, é altamente recomendado que a máquina possua o **Microsoft Word instalado e licenciado**, pois o sistema utiliza a integração COM nativa para uma renderização perfeita do documento.

---

## 4. Guia Passo a Passo das Ferramentas

### Menu Principal
Ao abrir o programa, você verá o painel principal. No canto superior direito, há o botão **"Modo Claro / Modo Escuro"** que altera a paleta de cores de todas as telas instantaneamente. Escolha o módulo desejado clicando em seu botão correspondente:

---

### 4.1 Gerador de Contrato EACE
Permite preencher e gerar um arquivo Word de contrato com tabelas detalhadas das escolas de forma automatizada com poucos cliques.

#### Fluxo de Operação:
1. **Aba "Arquivos":** 
   - O app **já vem configurado de fábrica** com a planilha permanente `EACENOC.xlsx` e os modelos de contrato Word. Eles são carregados de forma 100% automática.
   - Caso queira utilizar um modelo de contrato ou uma lista de escolas diferente, você pode clicar em **"Selecionar Template DOCX"** ou **"Selecionar Planilha EACE"** para carregar os arquivos manualmente de qualquer pasta.
2. **Aba "Contratada":** Preencha os campos com os dados cadastrais da empresa (Razão Social, CNPJ, Endereço, etc.).
3. **Aba "Suporte":** Forneça o e-mail e telefone de suporte ao cliente, bem como a URL ou nome do sistema de chamados.
4. **Aba "Assinaturas":** Informe o local, a data de assinatura (já preenchida com o dia atual) e os nomes dos signatários (Contratante, Testemunha, Contratada).
5. **Aba "INEPs":** 
   - No campo superior, digite os códigos INEP das escolas que farão parte do contrato. Você pode colar vários códigos separados por vírgula ou um por linha.
   - Clique em **"Adicionar INEPs"**.
   - O sistema buscará essas escolas na planilha configurada. Se encontrar mais de uma escola, abrirá uma tela de **Preenchimento em Lote** perguntando a Mensalidade (R$) e a data de instalação de cada uma. Se for apenas uma escola, abrirá uma tela simplificada.
   - Você pode preencher os valores um a um ou digitar na área de preenchimento rápido (lote) e clicar em **"Aplicar a Todos"**. Clique em **"Confirmar"**.
   - As escolas serão listadas na tabela abaixo com suas respectivas informações (INEP, Endereço, Latitude/Longitude, Velocidade da Banda, Mensalidade e Data de Instalação). Você pode remover qualquer item selecionando-o e clicando em **"Remover Selecionado"**.
6. **Aba "Gerar Contrato":** 
   - Clique em **"GERAR CONTRATO FINAL (WORD)"**.
   - Escolha o modelo de contrato desejado (ex: **Contrato BITNET** ou **Contrato ST1**). O aplicativo criará o documento, preencherá as cláusulas com os dados cadastrais, desenhará a tabela de escolas no local correto, calculará a soma total mensal dos serviços e abrirá um diálogo para você salvar o arquivo gerado em seu computador.

---

### 4.2 Consulta EACE & Verificador de Velocidade
Uma ferramenta indispensável para consultas imediatas durante chamados e validação técnica de velocidade de internet instalada na escola.

#### Fluxo de Operação:
1. Ao entrar no módulo, verifique a mensagem sob a barra de busca:
   - **Verde ("Cache local carregado"):** O aplicativo está pronto para fazer consultas rápidas instantâneas.
   - **Laranja ("Base local vazia"):** Clique no botão azul **"Atualizar da Nuvem"** para baixar a base de dados do Google Sheets pela primeira vez.
2. **Realizar Busca:** No campo **"Código INEP ou Nome"**, digite o número do INEP da escola ou partes do seu nome. Pressione `Enter` ou clique em **"BUSCAR DADOS"**.
   - Se houver apenas um resultado, a ficha da escola será carregada imediatamente.
   - Se houver múltiplos registros parecidos, uma janela popup exibirá a lista. Selecione a escola correta e clique em **"Visualizar Escola"**.
3. **Visualizar e Copiar:** Todas as informações cadastrais relevantes serão exibidas do lado esquerdo da tela. Para compartilhar rapidamente a ficha com um colega ou colar em um sistema de chamados, clique no botão verde **"📋 Copiar Ficha"**.
4. **Verificador de Velocidade (Painel Direito):**
   - O aplicativo mostra a velocidade contratada e a velocidade mínima requerida estipulada no edital para a escola selecionada.
   - Faça o teste de velocidade na escola, copie o valor medido de download e cole no campo **"Resultado do Speedtest"**.
   - Clique no botão roxo **"CALCULAR"** (ou pressione `Enter`).
   - O painel exibirá imediatamente o status: **APROVADO** (verde, se a velocidade medida for maior ou igual ao mínimo exigido) ou **REPROVADO** (vermelho, se estiver abaixo do mínimo exigido).
   - Uma barra de progresso visual com linha tracejada indicará qual a porcentagem da velocidade exigida foi atingida.

---

### 4.3 Unificador de Evidências Multiformato
Reúne arquivos soltos que documentam a instalação da escola em um PDF unificado e leve para envio de faturamento ou comprovação de instalação.

#### Recursos Suportados:
- **Imagens:** `.png`, `.jpg`, `.jpeg` (com correção automática de rotação e redimensionamento inteligente).
- **Textos:** `.txt` (converte o arquivo de texto em uma página visual de documento PDF).
- **Documentos:** `.pdf` existentes e arquivos Word `.docx` (convertidos para PDF dinamicamente).

#### Fluxo de Operação:
1. Clique em **"Alterar"** no topo da tela caso queira mudar a pasta onde o PDF final será salvo.
2. Clique no botão azul **"+ Adicionar"** para selecionar os arquivos de evidências do seu computador. Você pode selecionar múltiplos arquivos ao mesmo tempo.
3. Organize a sequência de páginas do PDF: selecione um arquivo na listagem e clique em **"Subir"** ou **"Descer"** para reordená-lo.
4. **Pré-visualização (Painel Direito):** Ao clicar em qualquer arquivo da lista, a miniatura dele aparecerá no painel de visualização lateral. Imagens e textos são renderizados em tela, enquanto PDFs e DOCXs exibem um ícone indicativo do formato.
5. Deixe marcada a caixa **"Otimizar tamanho das imagens"** para comprimir imagens pesadas, reduzindo drasticamente o tamanho final do arquivo PDF para que ele possa ser enviado facilmente por e-mail ou sistemas internos.
6. Preencha o campo **"Nome do PDF Final"** com a nomenclatura correta (ex: *Evidencias_INEP_12345678*).
7. Clique em **"GERAR PDF UNIFICADO AGORA"**. Um modal de progresso informará cada arquivo processado e confirmará o salvamento do PDF ao término.

---

### 4.4 Conversor Bidirecional PDF ⇄ Word
Uma ferramenta utilitária rápida para conversão em lote de arquivos de documentos de trabalho.

#### Fluxo de Operação:
1. No campo **"Sentido da Conversão"**, selecione o fluxo desejado:
   - **PDF para Word (.docx):** Transforma documentos fechados em arquivos editáveis.
   - **Word (.docx) para PDF:** Converte modelos de texto em arquivos finais PDF.
2. Clique no botão azul **"Alterar"** no topo para escolher onde salvar os novos arquivos.
3. Clique em **"+ Selecionar Arquivos"** para adicionar os arquivos a serem convertidos. O aplicativo aceita múltiplos documentos para execução simultânea (lote).
4. Clique no botão laranja **"EXECUTAR CONVERSÃO EM LOTE AGORA"**. O software converterá um a um em segundo plano sem congelar sua tela e exibirá um resumo de sucesso ao final do processo.

---

### 4.5 Dashboard de Monitoramento EACE
Fornece métricas, indicadores gerenciais e monitoramento completo de progresso do cronograma de instalações do projeto EACE.

#### Recursos Principais:
1. **Cards de Métricas:**
   - **Total de Escolas:** Número de escolas presentes no filtro aplicado.
   - **Instaladas / Concluídas:** Total de escolas cujo status indica conclusão (*OK, Instalado, Ativo, etc.*).
   - **Em Progresso:** Escolas com instalação em andamento ou pendência detectada.
   - **Taxa de Conclusão:** Porcentagem das instalações executadas em relação ao total.
2. **Filtros Avançados (Combinações em Tempo Real):**
   - **Regional / UF:** Selecione a regional técnica ou a Unidade Federativa. O sistema atualizará os demais filtros automaticamente para exibir apenas opções válidas.
   - **Município:** Selecione a cidade específica da escola.
   - **Status:** Filtre por status (ex: *Pendente, OK, Vistoria, etc.*).
   - **Busca:** Digite parte do INEP ou nome da escola para pesquisar dinamicamente na tabela.
   - **Mês/Ano:** Filtre pela data de instalação das escolas para verificar a produção de determinado mês (ex: *05/2026*).
   - **Ordenar Data:** Classifique a listagem de escolas de acordo com a data de instalação (do menor para o maior ou do maior para o menor).
3. **Exportação e Ações:**
   - Clique em **"📋 Copiar Registro Selecionado"** para transferir as informações completas da linha selecionada para a área de transferência.
   - Clique em **"📊 Copiar Tabela Filtrada"** para copiar toda a listagem que está na sua tela no formato tabulado do Excel. Em seguida, basta abrir uma planilha vazia no Excel e pressionar `Ctrl + V` para colar e gerar relatórios instantâneos.
4. **Sincronização:**
   - O dashboard utiliza o cache local `eace_dashboard_cache.json`.
   - Se for a primeira execução ou desejar obter os dados atualizados do projeto da nuvem, clique no botão **"Sincronizar Nuvem (Sheets)"**. O app conecta-se de forma assíncrona ao Google Sheets na aba dedicada a monitoramento e validação de instalação (GID `1113059829`), sincroniza os registros, atualiza os filtros locais e salva os dados no cache.

---

## 5. Solução de Problemas (Troubleshooting)

### ❓ Sincronizei com a nuvem, mas o sistema diz que trouxe zero resultados ou a tabela está vazia. O que houve?
**Causas prováveis:**
1. **Falta do arquivo `credentials.json`:** Sem as credenciais corretas do Google Cloud autorizadas na mesma pasta do executável, o programa não consegue baixar os dados de forma segura. Certifique-se de que o arquivo está no local de instalação (`%LOCALAPPDATA%\Programs\FeramentaEACEBitnet`).
2. **Bloqueio de Internet / Proxy:** Firewalls de redes corporativas podem bloquear as requisições às APIs do Google. Tente conectar-se a uma rede externa ou solicitar liberação de portas ao administrador da rede.
3. **Carga Inicial Offline:** Caso esteja operando de forma offline e na primeira execução o cache esteja vazio, certifique-se de que o arquivo **`Validação Instação - Plan1.csv`** está localizado na pasta do aplicativo. Ele serve como fallback automático para preencher o banco de dados local.

### ❓ O Unificador de Evidências ou o Conversor falhou ao tentar processar um arquivo DOCX (Word).
**Causas prováveis:**
- **Microsoft Word Fechado Incorretamente:** Se o Word estiver travado em segundo plano, a automação COM falhará. Abra o Gerenciador de Tarefas do Windows (`Ctrl + Shift + Esc`), finalize todos os processos chamados `Microsoft Word` ou `WINWORD.EXE` e tente novamente.
- **Word não Instalado:** Se a máquina não possuir o Microsoft Word, a conversão dependerá de bibliotecas python instaladas. Verifique se o pacote padrão possui suporte para a máquina local ou opte por realizar a conversão manual.

### ❓ O aplicativo abre, mas os botões parecem distorcidos ou as fontes estão ilegíveis.
- O aplicativo utiliza a fonte do sistema do Windows **"Segoe UI"** e o pacote **"Pillow"** para processamento gráfico. Certifique-se de que seu Windows não está com o zoom de tela/escala de DPI excessivamente alto (ex: 200%). Se necessário, ajuste a escala do monitor para 100% ou 125% nas configurações de exibição do Windows.

---
*Manual desenvolvido em conformidade com as regras operacionais Bitnet Telecom. Versão 6.1 (Junho/2026).* 