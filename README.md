# Assistente de Compras com IA

App desktop em Python para planejar compras de importação: junta o **saldo do WMS**, as **vendas** da planilha e os **pedidos em aberto**, projeta o estoque mês a mês, mostra onde vai faltar produto e pede para a IA (Claude) uma recomendação de compra.

## O problema

Em importação, o pedido de hoje só chega meses depois. Para não romper estoque, o comprador precisa cruzar três fontes (saldo no armazém, ritmo de venda e o que já está a caminho) e enxergar o saldo de cada SKU nos próximos meses. Isso era feito em planilha, SKU por SKU.

## O que ele faz

| Aba | Conteúdo |
|---|---|
| **Estoque & Necessidade** | Saldo disponível e reservado por SKU (via API do WMS), vendas do mês, projeção de venda e necessidade de compra |
| **Compras & Chegadas** | Pedidos em aberto com mês de chegada; importação e exportação em Excel |
| **Cenário** | Simula compras novas e vê o efeito no saldo de cada mês; gera um "cenário ideal" automático |
| **Visual Atual** | Grade colorida SKU x mês: verde com saldo, vermelho com ruptura |

Além disso:

- **Projeção de 12 meses** por SKU: saldo atual, menos venda projetada (global ou mês a mês), mais chegadas previstas.
- **Cenário ideal:** para itens de giro alto, o alvo é cobrir o mês de venda mais um de folga; para os demais, só cobrir as rupturas.
- **Recomendação por IA:** envia o resumo de cada SKU para o Claude, que devolve rupturas previstas, urgências, atenções e uma estratégia geral.
- **Exportações:** Excel, relatório HTML para PDF e comparativo orçado x realizado.

## Arquitetura

```
WMS (API Stokki) ── saldo por SKU ──┐
Google Sheets ───── vendas ─────────┼─► projeção 12 meses ─► grade / cenários ─► Excel / HTML
compras_data.json ─ pedidos ────────┘          │
                                               └─► Claude (Anthropic API) ─► recomendação de compra
```

## Como usar

```bash
pip install -r requirements.txt
cp .env.example .env               # token do WMS, planilha e chave da Anthropic
cp itens.exemplo.json itens.json   # troque pelos seus SKUs e categorias
python assistente_compras.py
```

Sem `ANTHROPIC_API_KEY` o app funciona normalmente e só o botão de IA fica desativado.

## Dados

- `itens.exemplo.json` traz um catálogo **fictício**. O seu catálogo vai em `itens.json`, que fica fora do Git.
- Pedidos, projeções e cenários ficam em `compras_data.json`, também fora do Git.
- A leitura da planilha usa a exportação CSV do Google Sheets, que exige compartilhamento por link. Use uma planilha sem dados sensíveis ou troque por acesso via conta de serviço.

## Stack

Python · CustomTkinter · API REST do WMS · Google Sheets · Anthropic API (Claude) · openpyxl
