# Robô STABIA / STUR - Regras da Reunião

Versão refatorada para fluxo lento, auditável e com funções mais diretas.

## Estratégias implementadas

1. **VCN com venda**
   - Se a coluna VCN vier com `Venda 1234` ou apenas um número de venda, busca pela coluna **Venda**.
   - Se retornar uma única venda, considera localizada.

2. **LATAM**
   - Se o estabelecimento contém `LATAM` e tem `*ABC123`, busca `ABC123` na coluna **Localizador**.
   - Compara o valor do Excel contra **Total Fornecedor** ou **Total Cliente**.

3. **Genérico / Hotel / Fornecedor**
   - Busca devagar combinando:
     - Fornecedor + Data de Emissão
     - Fornecedor Serviço + Data de Emissão
     - Fornecedor + Data de Início
     - Fornecedor Serviço + Data de Início
     - Fornecedor + Data de Término
     - Fornecedor Serviço + Data de Término
   - Se bater valor exato: escreve OK.
   - Se tiver candidatos próximos: escreve `POSSÍVEL VENDA`.
   - Se não encontrar: escreve `LANÇAMENTO MANUAL`.

## Como rodar

```bash
python3 src/main.py --arquivo input/excel_stur.xlsx
```

Ou sem parâmetro, pegando o Excel/CSV mais recente da pasta Downloads:

```bash
python3 src/main.py
```

## Observações

- O robô aguarda 3 segundos entre os passos principais para evitar consulta encavalada.
- A etapa final de alteração/pagamento do fornecedor ainda está segura/conservadora: apenas loga a venda validada. Quando houver fatura aberta para teste real, implementar os campos finais em `seguir_fluxo_venda_ok`.
- O parsing de Total Fornecedor/Total Cliente foi protegido para não confundir CNPJ/CPF com valor monetário.


## Ajuste desta versão

A busca genérica foi otimizada para não aplicar filtro de data quando a primeira busca por Fornecedor/Fornecedor Serviço não retorna nenhum resultado.

Fluxo atual para genéricos:

1. Limpa filtros.
2. Clica em Fornecedor ou Fornecedor Serviço.
3. Pesquisa o estabelecimento.
4. Se não vier resultado, não marca Manter Pesquisa e não pesquisa data.
5. Se vier resultado, aí sim marca Manter Pesquisa e testa Data de Emissão, Data de Início e Data de Término.


## Ajuste desta versão

Na busca genérica, o robô agora tenta a busca inicial em:

1. Fornecedor
2. Fornecedor Serviço
3. Localizador

Somente se uma dessas buscas iniciais retornar resultado, ele marca "Manter Pesquisa" e refina por:

- Data de Emissão
- Data de Início
- Data de Término

Se não houver resultado inicial em nenhuma das três colunas, ele marca como não localizado/manual e segue para o próximo item.


## Ajuste desta versão - regra do asterisco

Regra aplicada:

- Se o estabelecimento tiver `*` e for LATAM:
  - busca direto pelo código após `*` na coluna `Localizador`;
  - não tenta Fornecedor/Fornecedor Serviço.

- Se o estabelecimento tiver `*` e NÃO for LATAM:
  - busca por `Fornecedor` usando o estabelecimento completo;
  - busca por `Fornecedor Serviço` usando o estabelecimento completo;
  - busca por `Localizador` usando somente o código extraído após `*`.

- Se não tiver `*`:
  - busca apenas por `Fornecedor` e `Fornecedor Serviço`.

Em todos os casos genéricos, o robô só aplica filtro de data se a busca inicial retornar alguma linha.
