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
python3 src/main.py --arquivo input/2026-04-15_000000000455_d8266ce4caa4.xlsx
```

Ou sem parâmetro, pegando o Excel/CSV mais recente da pasta Downloads:

```bash
python3 src/main.py
```

## Observações

- O robô aguarda 3 segundos entre os passos principais para evitar consulta encavalada.
- A etapa final de alteração/pagamento do fornecedor ainda está segura/conservadora: apenas loga a venda validada. Quando houver fatura aberta para teste real, implementar os campos finais em `seguir_fluxo_venda_ok`.
- O parsing de Total Fornecedor/Total Cliente foi protegido para não confundir CNPJ/CPF com valor monetário.
