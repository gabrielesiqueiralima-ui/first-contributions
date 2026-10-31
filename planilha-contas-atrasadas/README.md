# Organizador de Contas Atrasadas + Plano de Quitação

Planilha (funciona no Google Sheets e no Excel) para ver quanto você deve, qual conta pagar primeiro e em quantos meses fica livre das dívidas.

- `Organizador_Contas_Atrasadas_EXEMPLO.xlsx` – preenchida com dados de exemplo.
- `Organizador_Contas_Atrasadas_EM_BRANCO.xlsx` – pronta para usar/vender.
- `gerar_planilha.py` – script que gera as duas versões (`python gerar_planilha.py`).

## Abas
| Aba | O que faz |
|---|---|
| **Painel** | Nome, renda e quanto pode pagar por mês (únicos campos). Total da dívida, já paguei, falta pagar, contas quitadas, progresso, "livre das dívidas em", dívida por tipo, gráfico pago × falta e lista **Pague Primeiro**. |
| **Contas** | Uma linha por conta: credor, tipo, vencimento, valor original, juros/multas, prioridade, negociação, observações. Calcula dias de atraso, total, já paguei, falta pagar e situação; nome fica riscado quando quitada. |
| **Pagamentos** | Data, conta (lista), valor e forma de pagamento. Cada lançamento é descontado da conta certa. |
| **Plano de Quitação** | Ordem de pagamento (prioridade+atraso ou menor dívida primeiro), mês em que começa e termina cada conta, data prevista e cronograma mês a mês com saldo e progresso. |
| **Como usar** | Passo a passo. |

Para usar no Google Sheets: Arquivo › Importar › Fazer upload.
