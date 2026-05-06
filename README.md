# Robô STUR - Base Inicial

Base pronta para você abrir no VS Code e começar os testes.

## O que já está preparado

- Leitura de arquivo `.xlsx`, `.xls` e `.csv`
- Extração do código da companhia aérea depois do `*`, com 6 dígitos/caracteres
- Login no STUR
- Navegação para `Operacional -> Vendas`
- Busca pelo localizador/código extraído
- Validação de venda encontrada/não encontrada
- Comparação entre valor do Excel e `Total do Fornecedor` retornado pelo site
- Escrita do resultado na última coluna da planilha
- Fluxo base para abrir venda, acessar pagamento do fornecedor, editar, salvar e gravar
- Logs em arquivo e terminal
- `.env` para configurações sensíveis
- Base para empacotar com PyInstaller

## Instalação

Crie o ambiente virtual:

```bash
python -m venv .venv
```

Ative:

Windows:

```bash
.venv\Scripts\activate
```

Mac/Linux:

```bash
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
playwright install chromium
```

A instalação do Playwright exige baixar os navegadores com `playwright install`. A própria documentação oficial informa esse passo para uso com Python.

## Configuração

Copie o arquivo de exemplo:

Windows:

```bash
copy .env.example .env
```

Mac/Linux:

```bash
cp .env.example .env
```

Preencha no `.env`:

```env
STUR_USER=seu_usuario
STUR_PASSWORD=sua_senha
```

## Rodar

Com navegador visível:

```bash
python src/main.py --arquivo input/2026-04-15_000000000455_d8266ce4caa4.xlsx
```

Sem abrir navegador:

```bash
python src/main.py --arquivo input/2026-04-15_000000000455_d8266ce4caa4.xlsx --headless
```

## Gerar executável na VM Windows

```bash
pip install pyinstaller
pyinstaller --onefile --name robo_stur src/main.py
```

O executável ficará em:

```text
dist/robo_stur.exe
```

O PyInstaller empacota o interpretador Python e dependências dentro do pacote/executável, então o cliente normalmente não precisa instalar Python para executar o `.exe`.

## Arquivos importantes

```text
src/main.py
src/stur_automation.py
src/excel_service.py
src/models.py
src/config.py
src/logger_config.py
src/email_service.py
```

## Onde trocar XPaths/selectors

Abra:

```text
src/stur_automation.py
```

Procure por:

```python
SELECTORS = {
```

Troque os seletores conforme você for estudando a tela/vídeo.

## Observação importante

O fluxo de e-mail ficou como stub/base porque você comentou que ainda não tem acesso ao e-mail. Quando tiver IMAP, Outlook, Gmail ou Microsoft Graph, você implementa em `src/email_service.py`.