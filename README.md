# Dashboard de Vendas de Celulares — 2024

Este projeto apresenta uma planilha completa em Excel (`dashboard_vendas.xlsx` / `vendas.xlsx`) com dados fictícios de vendas de smartphones ao longo do ano de 2024 (janeiro a outubro), incluindo consolidação por data de venda e análise mensal com gráfico de desempenho.

---

## Estrutura do Arquivo Excel

O arquivo Excel contém três abas principais:

1. **`Dados`**:
   - Contém **20 registros de vendas** de smartphones das marcas **Apple (iPhone)**, **Samsung**, **Xiaomi** e **Motorola**.
   - Colunas:
     - `Data de Venda`: Datas distribuídas entre janeiro e outubro de 2024 (`DD/MM/YYYY`).
     - `Modelo`: Nome comercial do aparelho.
     - `Valor Unitário`: Preço unitário em Reais (`R$`).
     - `Quantidade`: Unidades vendidas na transação.
     - `Valor Total`: Calculado via fórmula dinâmica (`=Quantidade * Valor Unitário`).

2. **`Analise`**:
   - Tabela dinâmica consolidando o **Total de Vendas** agrupado por **Data de Venda** em ordem cronológica (`=SUMIF(...)`).
   - Inclui linha de **Total Geral** (`R$ 160.048,00`).

3. **`Analise por Mes`**:
   - Consolidação mensal de janeiro a outubro de 2024 com:
     - **Mês**: Nome do mês em português.
     - **Total de Vendas (R$)**: Soma mensal calculada via `=SUMIFS(...)`.
     - **Qtd de Transações**: Total de vendas no mês calculado via `=COUNTIFS(...)`.
     - **Ticket Médio (R$)**: Razão entre o faturamento mensal e a quantidade de transações (`=IF(C2=0, 0, B2/C2)`).
     - **TOTAL GERAL**: Soma de `R$ 160.048,00`, `20` transações e ticket médio geral de `R$ 8.002,40`.
   - **Gráfico de Barras (`Vendas por Mês — 2024`)**: Posicionado abaixo da tabela (`A14:D34`), exibindo a evolução do faturamento mensal em Reais (`R$`).

---

## Como Reproduzir

Para recriar o arquivo Excel a partir dos scripts Python no ambiente:

```bash
pip install openpyxl
python3 create_excel.py
python3 add_monthly_sheet.py
python3 add_chart.py
```
