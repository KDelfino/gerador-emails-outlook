# Gerador e Enviador de E-mails via Outlook Automation

Aplicativo desktop desenvolvido em Python para automação de envio em massa e criação de rascunhos de e-mails personalizados através do **Microsoft Outlook Desktop**, utilizando dados de qualquer planilha **Excel (.xlsx, .xls)** ou arquivo **CSV**.

---

## Principais Recursos

- **Compatibilidade Ampla com Planilhas**:
  - Suporte a arquivos .xlsx, .xls e .csv.
  - Detecção inteligente da linha de cabeçalho (ignora metadados e linhas vazias).
  - Suporte a múltiplas abas de planilhas.
  - Seleção dinâmica da coluna de e-mail e descarte de colunas sem dados.

- **Modelos Dinâmicos com Tags**:
  - Inserção de tags no formato #coluna# tanto no assunto quanto no corpo do e-mail.
  - Grade responsiva de botões de tags para inserção rápida no texto com um clique.
  - Suporte a formatação com quebra de linha, links clicáveis e negrito (**texto**).

- **Condições e Segmentação por Grupos**:
  - Filtragem por qualquer coluna da planilha (ex: grupo, cargo, status, filial).
  - Checkboxes para selecionar quais grupos processar.
  - **Modo de Modelo Único**: Envia o mesmo e-mail para todos os grupos selecionados.
  - **Modo de Múltiplos Modelos por Grupo**: Permite configurar assunto, corpo e anexos exclusivos para cada grupo diferente em um único disparo.
  - Botão para replicar modelos entre grupos.

- **Gerenciamento de Anexos**:
  - Adição de múltiplos arquivos anexos por e-mail ou exclusivos por grupo.
  - Visualização de tamanho e caminho do anexo.

- **Dois Modos de Execução**:
  1. **Salvar como Rascunhos (Recomendado)**: Gera todos os e-mails na pasta *Rascunhos* do Outlook para conferência antes do envio.
  2. **Envio Direto**: Dispara os e-mails imediatamente com popup de confirmação de segurança.

- **Interface Gráfica**:
  - Interface construída em Tkinter.
  - Janela de pré-visualização em tempo real do primeiro destinatário de cada grupo.
  - Console de logs integrado com barra de progresso e botão para cancelar execução.

---

## Requisitos do Sistema

- **Sistema Operacional**: Windows 10 / 11
- **Microsoft Outlook Desktop**: Instalado e conectado a uma conta de e-mail ativa.

> **Nota:** Não é necessário configurar senhas, portas SMTP ou tokens de autenticação. O aplicativo utiliza a API oficial COM do Windows (win32com.client) para se comunicar diretamente com o Outlook configurado no computador.

---

## Como Executar

### 1. Executável Standalone (.exe)
Basta dar um duplo clique no arquivo:
`	ext
Gerador_Emails_Outlook.exe
`

### 2. A partir do Código-Fonte (Python)
Caso prefira executar pelo Python:

1. Instale as dependências:
   `ash
   pip install openpyxl pywin32
   `
2. Execute o aplicativo:
   `ash
   python app.py
   `
   *(Ou utilize o arquivo executar_app.bat)*

---

## Como Utilizar (Passo a Passo)

1. **Selecionar Planilha**: Clique em *Selecionar Planilha...* e escolha seu arquivo Excel ou CSV.
2. **Aba e Cabeçalho**: Caso o arquivo tenha várias abas, escolha a aba desejada. O cabeçalho é detectado automaticamente, mas pode ser ajustado se necessário.
3. **Coluna de E-mail**: Selecione a coluna que contém os endereços de e-mail dos destinatários.
4. **(Opcional) Condições / Grupos**:
   - Selecione uma coluna no campo *Condição / Grupo por Coluna* (ex: grupo).
   - Marque ou desmarque os grupos desejados.
   - Escolha se deseja usar o mesmo modelo para todos ou modelos diferentes por grupo.
5. **Montar o E-mail**:
   - Digite o assunto e o corpo do e-mail.
   - Clique nos botões de tags para inserir dados da planilha (ex: #nome#, #usuario#, #senha#).
6. **Anexos**: Clique em *Adicionar Arquivo(s)...* se desejar anexar documentos.
7. **Pré-visualizar**: Clique em *Pré-visualizar Exemplo* para checar o resultado final formatado.
8. **Executar**:
   - Escolha *Salvar como Rascunhos no Outlook* ou *Enviar Diretamente pelo Outlook*.
   - Clique no botão principal para iniciar o processamento.

---

## Como Gerar o Executável (.exe)

Para compilar novamente o executável a partir do código:

`ash
pip install pyinstaller
python -m PyInstaller --noconsole --onefile --name Gerador_Emails_Outlook app.py
`

---

## Tecnologias Utilizadas

- **Python 3.11+**
- **Tkinter / ttk**: Interface gráfica nativa.
- **openpyxl**: Manipulação e leitura de planilhas Excel.
- **pywin32 / pythoncom**: Automação nativa da API Microsoft Outlook (MAPI/COM).
- **PyInstaller**: Empacotamento do executável standalone.
