# Robô STABIA — Fluxo Financeiro / Conferências e Baixas

Esta versão substitui o fluxo antigo de **Operacional -> Vendas** pelo novo fluxo informado:

1. Acessar **Financeiro**.
2. Acessar **Conferências e Baixas**.
3. Acessar **Conferências e Baixas por Conferência**.
4. Na tabela, pesquisar usando a coluna **Descrição**.
5. Se encontrar resultado inicial, marcar **Manter Pesquisa** e pesquisar pela coluna **Dt.Conferência**.
6. Validar a coluna **Valor** ignorando o sinal positivo/negativo.
   - Exemplo: `-100,00` no STUR bate com `100,00` no Excel.
7. Quando encontrar o valor correto, clicar no primeiro ícone **Editar** da linha.
8. O fluxo para nesse ponto, aguardando as próximas orientações da tela de edição.

## Arquivos alterados/criados

- `src/main.py`: fluxo principal novo para Financeiro.
- `src/stur_financeiro_automation.py`: funções específicas da tela de Conferências e Baixas.
- `src/main_vendas_antigo_backup.py`: backup do fluxo antigo de Vendas.

## Observação técnica

A tabela de Conferências e Baixas possui `colspan` na coluna de ações. Por isso, o mapeamento de headers expande `colspan` antes de associar cada `td` com sua respectiva coluna.


## Múltiplos arquivos de entrada

Esta versão aceita uma ou duas planilhas/CSVs, ou até mais arquivos, se estiverem na pasta `input/`.

Formas de uso:

```bash
python3 src/main.py
```

Sem parâmetros, o robô processa todos os arquivos `.xlsx`, `.xls` e `.csv` que estiverem na pasta `input/`.

Também é possível informar arquivos manualmente:

```bash
python3 src/main.py --arquivo input/arquivo_antigo.xlsx --arquivo input/CLARA\ MES\ 6.csv
```

Ou informar uma pasta:

```bash
python3 src/main.py --pasta input
```

## Layout Clara

Para o CSV novo da Clara, o robô identifica automaticamente o layout quando encontra as colunas:

- `Transação`
- `Valor original`
- `Valor em R$`

Mapeamento usado:

- descrição para busca no STUR: coluna `Transação`;
- valor para comparação: coluna `Valor em R$`; se não existir, usa `Valor original`;
- data para `Dt.Conferência`: coluna `Data da Transação`.

A comparação de valor continua ignorando positivo/negativo. Exemplo: `-100,00` no STUR bate com `100,00` no arquivo.
