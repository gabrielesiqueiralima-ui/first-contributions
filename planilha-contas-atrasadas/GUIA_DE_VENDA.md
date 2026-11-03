# Guia para vender na Kiwify (checklist)

## O que está pronto nesta pasta
| Arquivo | Para quê |
|---|---|
| `Organizador_Contas_Atrasadas_EM_BRANCO.xlsx` | O produto. Fórmulas travadas e ocultas (Excel). |
| `Organizador_Contas_Atrasadas_EXEMPLO.xlsx` | Preenchida, para tirar prints para a página de vendas. |
| `proteger-google-sheets.gs` | Script que trava as fórmulas no Google Sheets com um clique. |
| `Como_acessar_sua_planilha.pdf` | Material para quem comprar (vai na área de membros). |

## 1. Montar a versão Google Sheets (uma vez)
1. Google Sheets › **Arquivo › Importar › Fazer upload** › `Organizador_Contas_Atrasadas_EM_BRANCO.xlsx` › **Substituir planilha** (ou criar nova).
2. Renomeie para **Organizador de Contas Atrasadas**.
3. **Extensões › Apps Script** › apague o conteúdo › cole `proteger-google-sheets.gs` › **Salvar**.
4. Escolha a função `protegerPlanilha` › **Executar** › autorize (Avançado › Acessar projeto).
5. Confira: tente apagar uma célula cinza da aba Contas — deve aparecer um aviso.
6. **Compartilhar** › Acesso geral: **Qualquer pessoa com o link** › **Leitor**.
7. Copie o link e troque o final por `/copy`:
   `https://docs.google.com/spreadsheets/d/SEU_ID/copy`
8. Abra esse link numa janela anônima com outra conta Google e confira se a cópia funciona.

> No Google Sheets não existe jeito de esconder as fórmulas de quem tem a cópia (a pessoa vira dona). A proteção evita que ela apague algo sem querer — é o padrão de quem vende planilha.

## 2. Cadastrar na Kiwify
1. Crie a conta e complete a verificação (documento e dados bancários).
2. **Produtos › Criar produto** › produto digital com **área de membros**.
3. Nome, descrição, preço, Pix e cartão.
4. Na área de membros, crie um módulo "Sua planilha" com:
   - o **link `/copy`** (botão/texto "Clique aqui para pegar sua planilha");
   - o PDF `Como_acessar_sua_planilha.pdf`;
   - o arquivo `Organizador_Contas_Atrasadas_EM_BRANCO.xlsx` (para quem usa Excel).
5. Garantia de **7 dias** (direito de arrependimento do CDC).
6. Faça uma compra de teste (ou use o modo de teste da Kiwify) e veja o acesso chegando.
7. Coloque o link do checkout na página de vendas e na bio do Instagram.

## 3. Página de vendas
- Use textos e prints **próprios** (tire prints da versão EXEMPLO no Google Sheets).
- Diga claramente: funciona no Google Sheets (grátis) e no Excel, celular e computador.

## Senha do Excel
A proteção do `.xlsx` tem senha, que **não** fica salva neste repositório. Para gerar de novo:
```bash
pip install openpyxl reportlab
PLANILHA_SENHA="sua-senha" python gerar_planilha.py
# recalcule os dois .xlsx (ex.: LibreOffice) para gravar os valores
PLANILHA_SENHA="sua-senha" python gerar_planilha.py --finalizar
CONTATO="seu@email.com" python gerar_pdf_acesso.py   # contato opcional no PDF
```
