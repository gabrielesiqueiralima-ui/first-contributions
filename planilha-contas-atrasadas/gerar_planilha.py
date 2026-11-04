"""Gera a planilha "Organizador de Contas Atrasadas" com Plano de Quitação.

Uso:
    python gerar_planilha.py            # gera a versão com exemplo e a em branco
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- estilo
FONT = "Arial"
VERDE = "2E6B5E"
BEGE = "F3E9D6"
VERDE_CLARO = "DCE6DC"
AZUL_CLARO = "DCE3EE"
ROSA = "F4D3CC"
AMARELO = "F7EBC4"
LILAS = "E6E2EE"
CINZA = "F3F3F0"
BRANCO = "FFFFFF"
INPUT = "FFF8EC"  # células que a pessoa preenche no Painel

MOEDA = '"R$" #,##0.00'
DATA = "dd/mm/yyyy"
MES_ANO = "mmm/yyyy"

fino = Side(style="thin", color="D9D6CE")
BORDA = Border(left=fino, right=fino, top=fino, bottom=fino)
grosso = Side(style="medium", color="BDB8AC")
BORDA_HEADER = Border(left=fino, right=fino, top=fino, bottom=grosso)

N_CONTAS = 50       # linhas de contas
N_PAG = 500         # linhas de pagamentos
N_MESES = 60        # meses no cronograma
C0, C1 = 6, 6 + N_CONTAS - 1          # linhas de dados em Contas
P0, P1 = 6, 6 + N_PAG - 1             # linhas de dados em Pagamentos

TIPOS = ["Cartão de crédito", "Empréstimo", "Financiamento", "Água", "Luz",
         "Internet/Telefone", "Aluguel/Condomínio", "Escola/Faculdade",
         "Impostos", "Saúde", "Outros"]
PRIORIDADES = ["Alta", "Média", "Baixa"]
NEGOCIACAO = ["Não negociei", "Tentando negociar", "Acordo fechado", "Sem acordo"]
FORMAS = ["Pix", "Boleto", "Cartão", "Transferência", "Débito automático", "Dinheiro"]
ESTRATEGIAS = ["Prioridade e atraso (recomendado)", "Menor dívida primeiro"]

PLANO = "'Plano de Quitação'"


def f(bold=False, size=10, color="333333", italic=False, strike=False, name=FONT):
    return Font(name=name, bold=bold, size=size, color=color, italic=italic, strike=strike)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def style(cell, *, font=None, bg=None, fmt=None, h="left", v="center", border=True, wrap=False):
    cell.font = font or f()
    if bg:
        cell.fill = fill(bg)
    if fmt:
        cell.number_format = fmt
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if border:
        cell.border = BORDA


def titulo(ws, rng, texto, bg=BEGE, size=14):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = texto
    c.font = Font(name="Courier New", bold=False, size=size, color="333333")
    c.fill = fill(bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[c.row].height = 34


def subtitulo(ws, rng, texto):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = texto
    c.font = f(italic=True, size=9, color="888888")
    c.alignment = Alignment(horizontal="center")


def header(ws, row, col0, nomes, bg=BEGE):
    for i, nome in enumerate(nomes):
        c = ws.cell(row=row, column=col0 + i, value=nome)
        c.font = Font(name="Courier New", bold=True, size=9, color="333333")
        c.fill = fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDA_HEADER
    ws.row_dimensions[row].height = 32


def lista(ws_listas, col, valores, titulo_):
    ws_listas.cell(row=1, column=col, value=titulo_).font = f(bold=True)
    for i, v in enumerate(valores, start=2):
        ws_listas.cell(row=i, column=col, value=v).font = f()
    letra = ws_listas.cell(row=1, column=col).column_letter
    return f"Listas!${letra}$2:${letra}${len(valores) + 1}"


def card(ws, label_rng, value_rng, label, formula, label_bg, fmt=None, value_bg=BRANCO, size=14):
    ws.merge_cells(label_rng)
    ws.merge_cells(value_rng)
    lc = ws[label_rng.split(":")[0]]
    vc = ws[value_rng.split(":")[0]]
    lc.value = label
    lc.font = Font(name="Courier New", size=8, color="444444")
    lc.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws[label_rng]:
        for c in row:
            c.fill = fill(label_bg)
            c.border = BORDA
    vc.value = formula
    vc.font = f(bold=True, size=size, color="222222")
    vc.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        vc.number_format = fmt
    for row in ws[value_rng]:
        for c in row:
            c.fill = fill(value_bg)
            c.border = BORDA


# ---------------------------------------------------------------- exemplo
HOJE = dt.date(2026, 10, 31)


def d(dias_atras):
    return HOJE - dt.timedelta(days=dias_atras)


EXEMPLO_CONTAS = [
    # conta, tipo, vencimento, valor original, juros/multas, prioridade, negociação, obs
    ("Cartão Visa", "Cartão de crédito", d(141), 2450.00, 380.00, "Alta", "Não negociei", "Pedir desconto para pagar à vista"),
    ("Conta de luz", "Luz", d(44), 287.40, 12.50, "Alta", "Não negociei", ""),
    ("Empréstimo pessoal", "Empréstimo", d(177), 3800.00, 640.00, "Média", "Acordo fechado", "Acordo em 10x de R$ 444"),
    ("Internet", "Internet/Telefone", d(70), 119.90, 8.20, "Baixa", "Não negociei", ""),
    ("Aluguel", "Aluguel/Condomínio", d(24), 1200.00, 60.00, "Alta", "Tentando negociar", "Conversar com o proprietário"),
    ("Mensalidade faculdade", "Escola/Faculdade", d(111), 890.00, 45.00, "Média", "Sem acordo", ""),
    ("Conta de água", "Água", d(31), 98.60, 4.10, "Alta", "Não negociei", ""),
    ("Financiamento do carro", "Financiamento", d(75), 1350.00, 210.00, "Alta", "Tentando negociar", "2 parcelas atrasadas"),
    ("IPVA 2026", "Impostos", d(182), 1180.00, 236.00, "Média", "Acordo fechado", "Parcelado em 4x no site da Sefaz"),
    ("Cartão Mastercard", "Cartão de crédito", d(99), 1640.00, 295.00, "Média", "Acordo fechado", ""),
    ("Empréstimo com a tia", "Outros", d(242), 500.00, 0.00, "Baixa", "Acordo fechado", "Pagar R$ 100 por mês"),
    ("Cheque especial", "Empréstimo", d(89), 760.00, 185.00, "Alta", "Não negociei", "Juros altíssimos, pagar logo"),
    ("Plano de saúde", "Saúde", d(39), 450.00, 22.50, "Alta", "Tentando negociar", "Ligar para a central"),
]

EXEMPLO_PAG = [
    (dt.date(2026, 10, 1), "Conta de luz", 299.90, "Pix"),
    (dt.date(2026, 10, 10), "Empréstimo pessoal", 444.00, "Boleto"),
    (dt.date(2026, 10, 12), "Cartão Visa", 500.00, "Pix"),
    (dt.date(2026, 10, 20), "Internet", 60.00, "Pix"),
    (dt.date(2026, 10, 3), "Conta de água", 102.70, "Pix"),
    (dt.date(2026, 10, 15), "Cartão Mastercard", 322.50, "Boleto"),
    (dt.date(2026, 9, 30), "Empréstimo com a tia", 100.00, "Pix"),
    (dt.date(2026, 10, 28), "Empréstimo com a tia", 100.00, "Pix"),
    (dt.date(2026, 10, 18), "Financiamento do carro", 780.00, "Boleto"),
    (dt.date(2026, 10, 25), "Aluguel", 600.00, "Transferência"),
    (dt.date(2026, 10, 26), "Plano de saúde", 236.25, "Pix"),
]


# ---------------------------------------------------------------- construção
def construir(com_exemplo: bool, caminho: str, dados=None):
    """dados (opcional): {"nome", "renda", "valor_mes", "contas": [...], "pagamentos": [...]}
    no mesmo formato de EXEMPLO_CONTAS / EXEMPLO_PAG, para gerar uma planilha já preenchida."""
    if dados is None and com_exemplo:
        dados = {"nome": "Ana", "renda": 4500, "valor_mes": 1200,
                 "contas": EXEMPLO_CONTAS, "pagamentos": EXEMPLO_PAG}
    dados = dados or {}
    lista_contas = dados.get("contas", [])
    lista_pag = dados.get("pagamentos", [])
    wb = Workbook()
    painel = wb.active
    painel.title = "Painel"
    contas = wb.create_sheet("Contas")
    pag = wb.create_sheet("Pagamentos")
    plano = wb.create_sheet("Plano de Quitação")
    como = wb.create_sheet("Como usar")
    listas = wb.create_sheet("Listas")

    for ws in (painel, contas, pag, plano, como):
        ws.sheet_view.showGridLines = False

    # ---------------- Listas (aba oculta com as opções dos menus)
    r_tipos = lista(listas, 1, TIPOS, "Tipos")
    r_prio = lista(listas, 2, PRIORIDADES, "Prioridade")
    r_neg = lista(listas, 3, NEGOCIACAO, "Negociação")
    r_formas = lista(listas, 4, FORMAS, "Forma de pagamento")
    r_estr = lista(listas, 5, ESTRATEGIAS, "Estratégia")
    for col in "ABCDE":
        listas.column_dimensions[col].width = 34
    listas.sheet_state = "veryHidden"  # não aparece em "Reexibir"

    # ================================================================ CONTAS
    ws = contas
    widths = {"A": 2, "B": 5, "C": 28, "D": 20, "E": 13, "F": 11, "G": 15, "H": 15, "I": 15,
              "J": 15, "K": 15, "L": 12, "M": 12, "N": 19, "O": 36, "Q": 12, "R": 12}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    titulo(ws, "B2:O2", "MINHAS CONTAS ATRASADAS")
    subtitulo(ws, "B3:O3", "Preencha as colunas brancas. As colunas cinzas calculam sozinhas — não precisa mexer nelas.")
    header(ws, 5, 2, ["Nº", "CONTA / CREDOR", "TIPO", "VENCIMENTO", "DIAS DE ATRASO",
                      "VALOR ORIGINAL", "JUROS E MULTAS", "VALOR TOTAL", "JÁ PAGUEI",
                      "FALTA PAGAR", "PRIORIDADE", "SITUAÇÃO", "NEGOCIAÇÃO", "OBSERVAÇÕES"])
    header(ws, 5, 17, ["ORDEM PAINEL", "ORDEM PLANO"], bg=CINZA)
    ws.freeze_panes = "D6"

    for r in range(C0, C1 + 1):
        i = r - C0
        ex = lista_contas[i] if i < len(lista_contas) else None
        ws[f"B{r}"] = f'=IF(C{r}="","",ROW()-{C0 - 1})'
        if ex:
            ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"], ws[f"G{r}"], ws[f"H{r}"] = ex[0], ex[1], ex[2], ex[3], ex[4]
            ws[f"L{r}"], ws[f"N{r}"], ws[f"O{r}"] = ex[5], ex[6], ex[7] or None
        ws[f"F{r}"] = f'=IF(OR(C{r}="",E{r}=""),"",MAX(0,TODAY()-E{r}))'
        ws[f"I{r}"] = f'=IF(C{r}="","",N(G{r})+N(H{r}))'
        ws[f"J{r}"] = f'=IF(C{r}="","",SUMIFS(Pagamentos!$D${P0}:$D${P1},Pagamentos!$C${P0}:$C${P1},C{r}))'
        ws[f"K{r}"] = f'=IF(C{r}="","",MAX(0,I{r}-J{r}))'
        ws[f"M{r}"] = f'=IF(C{r}="","",IF(AND(I{r}>0,K{r}<=0),"Quitada",IF(J{r}>0,"Pagando","Em aberto")))'
        # ordem do "Pague primeiro": prioridade (Alta=3, Média=2, Baixa=1) e depois dias de atraso
        ws[f"Q{r}"] = (f'=IF(OR(C{r}="",N(K{r})<=0),"",'
                       f'((L{r}="Alta")*3+(L{r}="Média")*2+(L{r}="Baixa")*1)*100000+N(F{r})+(1000-ROW())/10000)')
        # ordem do plano: depende da estratégia escolhida
        ws[f"R{r}"] = (f'=IF(Q{r}="","",IF({PLANO}!$J$4="{ESTRATEGIAS[1]}",'
                       f'10000000-K{r}+(1000-ROW())/1000000,Q{r}))')

        for col in "BCDEFGHIJKLMNO":
            c = ws[f"{col}{r}"]
            entrada = col in "CDEGHLNO"
            style(c, bg=BRANCO if entrada else CINZA,
                  h={"B": "center", "D": "center", "E": "center", "F": "center", "L": "center",
                     "M": "center", "N": "center"}.get(col, "right" if col in "GHIJK" else "left"))
        ws[f"E{r}"].number_format = DATA
        ws[f"F{r}"].number_format = '0 "dias"'
        for col in "GHIJK":
            ws[f"{col}{r}"].number_format = MOEDA
        for col in "QR":
            style(ws[f"{col}{r}"], font=f(size=8, color="999999"), bg=CINZA, fmt="0.00")
        ws.row_dimensions[r].height = 18

    ws.column_dimensions["Q"].hidden = True
    ws.column_dimensions["R"].hidden = True

    dv = DataValidation(type="list", formula1=r_tipos, allow_blank=True)
    dv.add(f"D{C0}:D{C1}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=r_prio, allow_blank=True)
    dv.add(f"L{C0}:L{C1}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=r_neg, allow_blank=True)
    dv.add(f"N{C0}:N{C1}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="greaterThan", formula1="1", allow_blank=True,
                        showErrorMessage=True, error="Digite uma data, ex.: 10/06/2026")
    dv.add(f"E{C0}:E{C1}")
    ws.add_data_validation(dv)

    rng = f"B{C0}:O{C1}"
    ws.conditional_formatting.add(
        f"C{C0}:C{C1}", FormulaRule(formula=[f'$M{C0}="Quitada"'], font=Font(strike=True, color="999999")))
    for txt, cor, fcor in (("Alta", ROSA, "9C3D2E"), ("Média", AMARELO, "8A6A10"), ("Baixa", AZUL_CLARO, "3A5378")):
        ws.conditional_formatting.add(f"L{C0}:L{C1}", FormulaRule(
            formula=[f'$L{C0}="{txt}"'], fill=fill(cor), font=Font(color=fcor)))
    for txt, cor, fcor in (("Quitada", VERDE_CLARO, "2E6B3E"), ("Pagando", AMARELO, "8A6A10"),
                           ("Em aberto", ROSA, "9C3D2E")):
        ws.conditional_formatting.add(f"M{C0}:M{C1}", FormulaRule(
            formula=[f'$M{C0}="{txt}"'], fill=fill(cor), font=Font(bold=True, color=fcor)))

    ws["C5"].comment = Comment("Escreva o nome do jeito que você reconhece (ex.: Cartão Visa). "
                               "Esse nome aparece na lista da aba Pagamentos.", "Organizador")
    ws["H5"].comment = Comment("Juros, multas e encargos que já foram cobrados. Se não souber, deixe 0 "
                               "e atualize quando receber o boleto.", "Organizador")
    ws["L5"].comment = Comment("Alta = essencial ou juros altos (luz, água, aluguel, cheque especial, cartão). "
                               "Média = importante. Baixa = pode esperar.", "Organizador")

    # ================================================================ PAGAMENTOS
    ws = pag
    for k, v in {"A": 2, "B": 14, "C": 30, "D": 16, "E": 20, "F": 36}.items():
        ws.column_dimensions[k].width = v
    titulo(ws, "B2:F2", "REGISTRO DE PAGAMENTOS", bg=AZUL_CLARO)
    subtitulo(ws, "B3:F3", "Cada pagamento lançado aqui é descontado automaticamente da conta certa.")
    header(ws, 5, 2, ["DATA", "CONTA", "VALOR PAGO", "FORMA DE PAGAMENTO", "OBSERVAÇÃO"], bg=AZUL_CLARO)
    ws.freeze_panes = "B6"
    for r in range(P0, P1 + 1):
        i = r - P0
        if i < len(lista_pag):
            e = lista_pag[i]
            ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"] = e
        for col in "BCDEF":
            style(ws[f"{col}{r}"], h="center" if col in "BE" else ("right" if col == "D" else "left"))
        ws[f"B{r}"].number_format = DATA
        ws[f"D{r}"].number_format = MOEDA
    dv = DataValidation(type="list", formula1=f"Contas!$C${C0}:$C${C1}", allow_blank=True,
                        showErrorMessage=True, error="Escolha uma conta da lista (cadastre antes na aba Contas).")
    dv.add(f"C{P0}:C{P1}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=r_formas, allow_blank=True)
    dv.add(f"E{P0}:E{P1}")
    ws.add_data_validation(dv)

    # ================================================================ PAINEL
    ws = painel
    for k, v in {"A": 2, "B": 24, "C": 12, "D": 16, "E": 16, "F": 12, "G": 3, "H": 20, "I": 18,
                 "J": 3, "K": 8}.items():
        ws.column_dimensions[k].width = v
    titulo(ws, "B2:I2", "ORGANIZADOR DE CONTAS ATRASADAS", bg=VERDE_CLARO)
    ws.merge_cells("B3:I3")
    ws["B3"] = '=IF(B6="","Olá! Vamos sair das dívidas, um passo de cada vez.","Olá, "&B6&"! Vamos sair das dívidas, um passo de cada vez.")'
    ws["B3"].font = f(italic=True, size=9, color="777777")
    ws["B3"].alignment = Alignment(horizontal="center")

    card(ws, "B5:C5", "B6:C6", "✎ SEU NOME", dados.get("nome"), LILAS, value_bg=INPUT)
    card(ws, "E5:F5", "E6:F6", "✎ RENDA MENSAL", dados.get("renda"), LILAS, MOEDA, value_bg=INPUT)
    card(ws, "H5:I5", "H6:I6", "✎ POSSO PAGAR POR MÊS", dados.get("valor_mes"), LILAS, MOEDA, value_bg=INPUT)
    ws["H5"].comment = Comment("Quanto sobra por mês, depois das despesas da casa, para pagar dívidas. "
                               "É esse valor que monta o Plano de Quitação.", "Organizador")

    card(ws, "B8:C8", "B9:C9", "TOTAL DA DÍVIDA", f"=SUM(Contas!I{C0}:I{C1})", ROSA, MOEDA)
    card(ws, "E8:F8", "E9:F9", "JÁ PAGUEI", f"=SUM(Contas!J{C0}:J{C1})", VERDE_CLARO, MOEDA)
    card(ws, "H8:I8", "H9:I9", "FALTA PAGAR", f"=SUM(Contas!K{C0}:K{C1})", BEGE, MOEDA)

    card(ws, "B11:C11", "B12:C12", "CONTAS QUITADAS",
         f'=COUNTIF(Contas!M{C0}:M{C1},"Quitada")&" de "&COUNTA(Contas!C{C0}:C{C1})', AZUL_CLARO)
    card(ws, "E11:F11", "E12:F12", "PROGRESSO", "=IF(B9=0,0,MIN(1,E9/B9))", "D5E8E4", "0%")
    card(ws, "H11:I11", "H12:I12", "LIVRE DAS DÍVIDAS EM",
         '=IF(B9=0,"—",IF(H9<=0,"Livre! 🎉",IF(N(H6)<=0,"Preencha o valor/mês",'
         'ROUNDUP(H9/H6,0)&IF(ROUNDUP(H9/H6,0)=1," mês"," meses"))))', LILAS)

    ws.merge_cells("E13:F13")
    ws["E13"] = '=REPT("▰",ROUND(E12*12,0))&REPT("▱",12-ROUND(E12*12,0))'
    ws["E13"].font = f(size=10, color=VERDE)
    ws["E13"].alignment = Alignment(horizontal="center")
    ws.merge_cells("H13:I13")
    ws["H13"] = (f'=IF(OR(N(H6)<=0,H9<=0),"","pagando "&TEXT(H6,"R$ #,##0")&" por mês · até "&'
                 f'TEXT({PLANO}!J6,"mm/yyyy"))')
    ws["H13"].font = f(size=8, italic=True, color="777777")
    ws["H13"].alignment = Alignment(horizontal="center")

    # Dívida por tipo
    ws.merge_cells("B15:F15")
    ws["B15"] = "DÍVIDA POR TIPO"
    header(ws, 15, 2, ["DÍVIDA POR TIPO"])
    for col in "CDEF":
        ws[f"{col}15"].fill = fill(BEGE)
        ws[f"{col}15"].border = BORDA_HEADER
    header(ws, 16, 2, ["TIPO", "CONTAS", "TOTAL", "FALTA PAGAR", "% DO QUE FALTA"])
    ws.row_dimensions[15].height = 22
    t0 = 17
    for i, tipo in enumerate(TIPOS):
        r = t0 + i
        ws[f"B{r}"] = tipo
        ws[f"C{r}"] = f'=COUNTIF(Contas!$D${C0}:$D${C1},B{r})'
        ws[f"D{r}"] = f'=SUMIFS(Contas!$I${C0}:$I${C1},Contas!$D${C0}:$D${C1},B{r})'
        ws[f"E{r}"] = f'=SUMIFS(Contas!$K${C0}:$K${C1},Contas!$D${C0}:$D${C1},B{r})'
        ws[f"F{r}"] = f'=IF($H$9=0,0,E{r}/$H$9)'
        for col in "BCDEF":
            style(ws[f"{col}{r}"], font=f(size=9), h="left" if col == "B" else "center")
        ws[f"D{r}"].number_format = MOEDA
        ws[f"E{r}"].number_format = MOEDA
        ws[f"F{r}"].number_format = "0%"
    rt = t0 + len(TIPOS)
    ws[f"B{rt}"] = "TOTAL"
    ws[f"C{rt}"] = f"=SUM(C{t0}:C{rt - 1})"
    ws[f"D{rt}"] = f"=SUM(D{t0}:D{rt - 1})"
    ws[f"E{rt}"] = f"=SUM(E{t0}:E{rt - 1})"
    ws[f"F{rt}"] = f"=SUM(F{t0}:F{rt - 1})"
    for col in "BCDEF":
        style(ws[f"{col}{rt}"], font=f(bold=True, size=9), bg=BEGE, h="left" if col == "B" else "center")
    ws[f"D{rt}"].number_format = MOEDA
    ws[f"E{rt}"].number_format = MOEDA
    ws[f"F{rt}"].number_format = "0%"

    # Pago x Falta pagar (gráfico)
    ws.merge_cells("H15:I15")
    header(ws, 15, 8, ["PAGO × FALTA PAGAR"], bg=VERDE_CLARO)
    ws["I15"].fill = fill(VERDE_CLARO)
    ws["H26"], ws["I26"] = "Já paguei", "=E9"
    ws["H27"], ws["I27"] = "Falta pagar", "=H9"
    for r in (26, 27):
        style(ws[f"H{r}"], font=f(size=9))
        style(ws[f"I{r}"], font=f(size=9), fmt=MOEDA, h="right")
    ch = DoughnutChart()
    ch.holeSize = 55
    ch.add_data(Reference(ws, min_col=9, min_row=26, max_row=27), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=8, min_row=26, max_row=27))
    s = ch.series[0]
    for idx, cor in ((0, "8DB08A"), (1, "E3B3A3")):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = cor
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        s.dPt.append(pt)
    s.dLbls = DataLabelList()
    s.dLbls.showPercent = True
    s.dLbls.showVal = False
    s.dLbls.showCatName = False
    s.dLbls.showSerName = False
    ch.legend.position = "b"
    ch.width, ch.height = 9.2, 5.6
    ws.add_chart(ch, "H16")

    # Pague primeiro
    pp = rt + 3
    ws.merge_cells(f"B{pp}:I{pp}")
    header(ws, pp, 2, ["PAGUE PRIMEIRO — SUAS PRIORIDADES"], bg=ROSA)
    for col in "CDEFGHI":
        ws[f"{col}{pp}"].fill = fill(ROSA)
        ws[f"{col}{pp}"].border = BORDA_HEADER
    ws.row_dimensions[pp].height = 22
    header(ws, pp + 1, 2, ["CONTA", "", "TIPO", "DIAS DE ATRASO", "PRIORIDADE", "", "SITUAÇÃO", "FALTA PAGAR"])
    ws.merge_cells(f"B{pp + 1}:C{pp + 1}")
    ws.merge_cells(f"F{pp + 1}:G{pp + 1}")
    q0 = pp + 2
    for k in range(1, 11):
        r = q0 + k - 1
        ws[f"K{r}"] = (f'=IFERROR(MATCH(LARGE(Contas!$Q${C0}:$Q${C1},{k}),Contas!$Q${C0}:$Q${C1},0),"")')
        ws.merge_cells(f"B{r}:C{r}")
        ws.merge_cells(f"F{r}:G{r}")
        ws[f"B{r}"] = f'=IF($K{r}="",IF({k}=1,"Nenhuma conta em aberto 🎉",""),INDEX(Contas!$C${C0}:$C${C1},$K{r}))'
        ws[f"D{r}"] = f'=IF($K{r}="","",INDEX(Contas!$D${C0}:$D${C1},$K{r}))'
        ws[f"E{r}"] = f'=IF($K{r}="","",INDEX(Contas!$F${C0}:$F${C1},$K{r}))'
        ws[f"F{r}"] = f'=IF($K{r}="","",INDEX(Contas!$L${C0}:$L${C1},$K{r}))'
        ws[f"H{r}"] = f'=IF($K{r}="","",INDEX(Contas!$M${C0}:$M${C1},$K{r}))'
        ws[f"I{r}"] = f'=IF($K{r}="","",INDEX(Contas!$K${C0}:$K${C1},$K{r}))'
        for col in "BCDEFGHI":
            style(ws[f"{col}{r}"], font=f(size=9), h="left" if col == "B" else "center")
        ws[f"E{r}"].number_format = '0 "dias"'
        ws[f"I{r}"].number_format = MOEDA
        ws[f"I{r}"].alignment = Alignment(horizontal="right")
    for txt, cor, fcor in (("Alta", ROSA, "9C3D2E"), ("Média", AMARELO, "8A6A10"), ("Baixa", AZUL_CLARO, "3A5378")):
        ws.conditional_formatting.add(f"F{q0}:G{q0 + 9}", FormulaRule(
            formula=[f'$F{q0}="{txt}"'], fill=fill(cor), font=Font(color=fcor)))
    ws.column_dimensions["K"].hidden = True
    ws.merge_cells(f"B{q0 + 10}:I{q0 + 10}")
    ws[f"B{q0 + 10}"] = ("A lista se organiza sozinha: primeiro prioridade Alta, depois Média e Baixa; "
                         "dentro de cada uma, quem está há mais tempo atrasada vem antes.")
    ws[f"B{q0 + 10}"].font = f(italic=True, size=8, color="888888")

    # ================================================================ PLANO DE QUITAÇÃO
    ws = plano
    for k, v in {"A": 2, "B": 7, "C": 26, "D": 18, "E": 11, "F": 15, "G": 15, "H": 12, "I": 12,
                 "J": 14, "K": 4, "L": 3, "M": 7, "N": 12, "O": 15, "P": 15, "Q": 15, "R": 44}.items():
        ws.column_dimensions[k].width = v
    titulo(ws, "B2:R2", "PLANO DE QUITAÇÃO", bg=VERDE_CLARO)
    subtitulo(ws, "B3:R3", "Mostra em que ordem pagar e em que mês cada conta fica quitada, usando o valor que você pode pagar por mês.")

    card(ws, "B4:C4", "B5:C5", "VALOR POR MÊS (do Painel)", "=N(Painel!H6)", LILAS, MOEDA, size=12)
    card(ws, "D4:F4", "D5:F5", "✎ COMEÇAR EM", "=DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)", LILAS, MES_ANO, value_bg=INPUT, size=12)
    ws.merge_cells("G4:I4")
    ws["G4"] = "✎ ESTRATÉGIA"
    ws["G4"].font = Font(name="Courier New", size=8, color="444444")
    ws["G4"].alignment = Alignment(horizontal="center", vertical="center")
    for col in "GHI":
        ws[f"{col}4"].fill = fill(LILAS)
        ws[f"{col}4"].border = BORDA
    # célula da estratégia: J4 (referenciada pela aba Contas). G5:I5 mostra explicação.
    ws["J4"] = ESTRATEGIAS[0]
    style(ws["J4"], font=f(bold=True, size=9), bg=INPUT, h="center", wrap=True)
    ws.merge_cells("J4:K5")
    ws.merge_cells("G5:I5")
    ws["G5"] = (f'=IF(J4="{ESTRATEGIAS[1]}","Quita primeiro as menores: você vê contas saindo da lista mais rápido.",'
                f'"Essenciais e mais atrasadas primeiro (igual ao Pague Primeiro).")')
    style(ws["G5"], font=f(size=8, italic=True, color="555555"), h="center", wrap=True)
    for col in "HI":
        ws[f"{col}5"].border = BORDA
    ws.row_dimensions[5].height = 30
    dv = DataValidation(type="list", formula1=r_estr, allow_blank=False)
    dv.add("J4")
    ws.add_data_validation(dv)
    ws["J4"].comment = Comment("Escolha a estratégia do plano na lista.", "Organizador")

    # resumo
    card(ws, "B6:C6", "B7:C7", "TOTAL A QUITAR", "=SUM(F11:F60)", ROSA, MOEDA, size=12)
    card(ws, "D6:F6", "D7:F7", "MESES ATÉ QUITAR TUDO",
         '=IF(B7<=0,0,IF(B5<=0,"Preencha o valor/mês",ROUNDUP(B7/B5,0)))', BEGE, '0 "meses"', size=12)
    card(ws, "G6:I6", "G7:I7", "ÚLTIMA PARCELA EM",
         '=IF(OR(B7<=0,B5<=0),"—",EDATE(D5,D7-1))', VERDE_CLARO, MES_ANO, size=12)
    ws["J6"] = '=IF(OR(B7<=0,B5<=0),"",EDATE(D5,D7-1))'   # usado pelo Painel
    ws["J6"].number_format = MES_ANO
    ws["J6"].font = f(size=8, color="FFFFFF")

    # ordem de pagamento
    header(ws, 10, 2, ["ORDEM", "CONTA", "TIPO", "PRIORI-DADE", "FALTA PAGAR", "ACUMULADO",
                       "COMEÇA NO MÊS", "QUITA NO MÊS", "QUITADA EM"], bg=BEGE)
    ws.merge_cells("B9:J9")
    ws["B9"] = "ORDEM DE PAGAMENTO — uma conta de cada vez, sem espalhar o dinheiro"
    ws["B9"].font = f(bold=True, size=10, color=VERDE)
    for k in range(1, N_CONTAS + 1):
        r = 10 + k
        ws[f"K{r}"] = f'=IFERROR(MATCH(LARGE(Contas!$R${C0}:$R${C1},{k}),Contas!$R${C0}:$R${C1},0),"")'
        ws[f"B{r}"] = f'=IF($K{r}="","",{k})'
        ws[f"C{r}"] = f'=IF($K{r}="","",INDEX(Contas!$C${C0}:$C${C1},$K{r}))'
        ws[f"D{r}"] = f'=IF($K{r}="","",INDEX(Contas!$D${C0}:$D${C1},$K{r}))'
        ws[f"E{r}"] = f'=IF($K{r}="","",INDEX(Contas!$L${C0}:$L${C1},$K{r}))'
        ws[f"F{r}"] = f'=IF($K{r}="","",INDEX(Contas!$K${C0}:$K${C1},$K{r}))'
        ws[f"G{r}"] = f'=IF($K{r}="","",SUM($F$11:F{r}))'
        ws[f"H{r}"] = f'=IF(OR($K{r}="",$B$5<=0),"",INT((G{r}-F{r})/$B$5)+1)'
        ws[f"I{r}"] = f'=IF(OR($K{r}="",$B$5<=0),"",ROUNDUP(G{r}/$B$5,0))'
        ws[f"J{r}"] = f'=IF(I{r}="","",EDATE($D$5,I{r}-1))'
        for col in "BCDEFGHIJ":
            style(ws[f"{col}{r}"], font=f(size=9), bg=CINZA if False else BRANCO,
                  h="left" if col in "CD" else ("right" if col in "FG" else "center"))
        ws[f"F{r}"].number_format = MOEDA
        ws[f"G{r}"].number_format = MOEDA
        ws[f"H{r}"].number_format = '"mês "0'
        ws[f"I{r}"].number_format = '"mês "0'
        ws[f"J{r}"].number_format = MES_ANO
        ws[f"K{r}"].font = f(size=8, color="FFFFFF")
    ws.column_dimensions["K"].hidden = False
    for txt, cor, fcor in (("Alta", ROSA, "9C3D2E"), ("Média", AMARELO, "8A6A10"), ("Baixa", AZUL_CLARO, "3A5378")):
        ws.conditional_formatting.add("E11:E60", FormulaRule(
            formula=[f'$E11="{txt}"'], fill=fill(cor), font=Font(color=fcor)))

    # cronograma mensal
    ws.merge_cells("M9:R9")
    ws["M9"] = "CRONOGRAMA MÊS A MÊS"
    ws["M9"].font = f(bold=True, size=10, color=VERDE)
    header(ws, 10, 13, ["MÊS", "MÊS/ANO", "PAGAR NO MÊS", "SALDO DEPOIS", "PROGRESSO", "ONDE VAI O DINHEIRO"],
           bg=VERDE_CLARO)
    for m in range(1, N_MESES + 1):
        r = 10 + m
        ativo = f'OR(NOT(ISNUMBER($D$7)),{m}>$D$7)'
        ws[f"M{r}"] = f'=IF({ativo},"",{m})'
        ws[f"N{r}"] = f'=IF(M{r}="","",EDATE($D$5,{m}-1))'
        ws[f"O{r}"] = f'=IF(M{r}="","",MIN($B$5,$B$7-({m}-1)*$B$5))'
        ws[f"P{r}"] = f'=IF(M{r}="","",MAX(0,$B$7-{m}*$B$5))'
        ws[f"Q{r}"] = (f'=IF(M{r}="","",IF(Painel!$B$9=0,0,(Painel!$B$9-P{r})/Painel!$B$9))')
        # conta onde o dinheiro começa a ir neste mês (+ quantas outras)
        a = f'(COUNTIF($G$11:$G$60,"<="&({m}-1)*$B$5)+1)'
        b = f'MIN(COUNT($G$11:$G$60),COUNTIF($G$11:$G$60,"<"&{m}*$B$5)+1)'
        ws[f"R{r}"] = (f'=IF(M{r}="","",INDEX($C$11:$C$60,{a})&IF({b}>{a}," + "&({b}-{a})&'
                       f'IF({b}-{a}=1," outra"," outras"),"")&IF(COUNTIF($I$11:$I$60,{m})>0,'
                       f'"  ✔ quita "&COUNTIF($I$11:$I$60,{m}),""))')
        for col in "MNOPQR":
            style(ws[f"{col}{r}"], font=f(size=9), h="left" if col == "R" else ("right" if col in "OP" else "center"))
        ws[f"N{r}"].number_format = MES_ANO
        ws[f"O{r}"].number_format = MOEDA
        ws[f"P{r}"].number_format = MOEDA
        ws[f"Q{r}"].number_format = "0%"
    ws.conditional_formatting.add(f"M11:R{10 + N_MESES}", FormulaRule(
        formula=['AND($M11<>"",COUNTIF($I$11:$I$60,$M11)>0)'], fill=fill("EEF4EC")))
    ws.merge_cells(f"M{11 + N_MESES}:R{11 + N_MESES}")
    ws[f"M{11 + N_MESES}"] = ('=IF(AND(ISNUMBER(D7),D7>' + str(N_MESES) + '),"O plano passa de '
                              + str(N_MESES) + ' meses: tente aumentar o valor por mês ou negociar descontos.","")')
    ws[f"M{11 + N_MESES}"].font = f(italic=True, size=9, color="9C3D2E")
    ws.freeze_panes = "B11"

    ws.merge_cells("B62:J65")
    ws["B62"] = ("Como o plano funciona: todo o valor do mês vai para a 1ª conta da ordem até ela ser quitada; "
                 "o que sobrar no mês já passa para a próxima. Ele usa o 'Falta pagar' de hoje e não prevê juros "
                 "futuros — por isso, quanto antes você negociar e pagar, melhor. Quando você lança pagamentos "
                 "ou muda o valor por mês, o plano se refaz sozinho. Dica: se fechar um acordo com parcela fixa, "
                 "lance o valor total do acordo em 'Valor original' e zere os juros.")
    ws["B62"].font = f(size=9, color="555555")
    ws["B62"].alignment = Alignment(wrap_text=True, vertical="top")

    # ================================================================ COMO USAR
    ws = como
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 100
    titulo(ws, "B2:C2", "COMO USAR", bg=BEGE)
    passos = [
        ("1", "Na aba Painel, preencha só as 3 células claras: seu nome, sua renda mensal e quanto você pode pagar por mês."),
        ("2", "Na aba Contas, anote cada conta atrasada uma vez: nome, tipo, vencimento, valor original e juros/multas "
              "(se souber). Escolha a prioridade: Alta para o essencial (luz, água, aluguel) ou para juros muito altos "
              "(cartão, cheque especial)."),
        ("3", "Sempre que pagar algo, lance na aba Pagamentos: data, conta (escolha na lista), valor e forma de pagamento. "
              "Pode lançar vários pagamentos para a mesma conta."),
        ("4", "Pronto! As colunas cinzas calculam sozinhas: dias de atraso, total, quanto já pagou, quanto falta e a "
              "situação. Quando a conta é quitada, o nome fica riscado."),
        ("5", "No Painel você vê o total da dívida, o progresso, em quantos meses fica livre e a lista Pague Primeiro."),
        ("6", "Na aba Plano de Quitação você vê a ordem de pagamento e o mês em que cada conta será quitada. "
              "Escolha o mês de início e a estratégia: 'Prioridade e atraso' (recomendado) ou 'Menor dívida primeiro'."),
        ("★", "Células brancas/creme = você preenche. Células cinzas = fórmulas, não precisa mexer."),
        ("★", "Para começar do zero, apague os exemplos: nas abas Contas e Pagamentos, apague só as células brancas "
              "(não as cinzas)."),
        ("★", "As fórmulas são protegidas para não serem apagadas sem querer: só as células brancas/creme "
              "aceitam digitação. Se aparecer um aviso ao editar, é porque aquela célula calcula sozinha."),
        ("★", "Google Sheets: Arquivo › Importar › Fazer upload deste arquivo. Funciona no celular e no computador."),
        ("★", "Negociando? Use a coluna Negociação e as Observações para lembrar o que foi combinado com cada credor."),
    ]
    for i, (n, txt) in enumerate(passos):
        r = 4 + i
        ws[f"B{r}"] = n
        ws[f"C{r}"] = txt
        style(ws[f"B{r}"], font=f(bold=True, size=12, color=VERDE), bg=VERDE_CLARO if n != "★" else BEGE, h="center")
        style(ws[f"C{r}"], font=f(size=10), wrap=True)
        ws.row_dimensions[r].height = 36

    # impressão: paisagem, cabendo na largura da página
    for w in (painel, contas, pag, plano, como):
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True

    # proteção: só as células de preenchimento ficam liberadas; fórmulas travadas e ocultas
    entradas = {
        painel: ["B6:C6", "E6:F6", "H6:I6"],
        contas: [f"C{C0}:E{C1}", f"G{C0}:H{C1}", f"L{C0}:L{C1}", f"N{C0}:O{C1}"],
        pag: [f"B{P0}:F{P1}"],
        plano: ["D5:F5", "J4:K5"],
        como: [],
    }
    if dados.get("extra"):  # abas extras (ex.: agenda pessoal) antes de proteger
        dados["extra"](wb)
    proteger(wb, entradas, os.environ.get("PLANILHA_SENHA") or None)

    # ordem das abas e aba ativa
    wb.active = 0
    wb.save(caminho)


def proteger(wb, entradas, senha):
    """Trava todas as abas, deixando livres só as células de entrada.

    Células com fórmula ficam ocultas (a fórmula não aparece na barra). A senha vem da
    variável de ambiente PLANILHA_SENHA e não fica gravada neste script.
    """
    travada = Protection(locked=True, hidden=True)
    livre = Protection(locked=False, hidden=False)
    for ws, faixas in entradas.items():
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.protection = travada
        for faixa in faixas:
            for row in ws[faixa]:
                for c in row:
                    c.protection = livre
        p = ws.protection
        p.sheet = True
        if senha:
            p.password = senha
        # permite selecionar, mexer na largura das colunas e usar os menus de lista
        p.selectLockedCells = False
        p.selectUnlockedCells = False
        p.formatColumns = False
        p.formatRows = False
    listas = wb["Listas"]
    listas.protection.sheet = True
    if senha:
        listas.protection.password = senha
    wb.security = WorkbookProtection(lockStructure=True)
    if senha:
        wb.security.workbookPassword = senha


def finalizar(caminho, senha):
    """Reaplica a trava da estrutura e oculta de vez a aba Listas.

    O recálculo pelo LibreOffice (que grava os valores das fórmulas no arquivo) apaga a
    proteção da pasta de trabalho; aqui ela é recolocada direto no XML, sem perder os valores.
    """
    import re
    import shutil
    import tempfile
    import zipfile
    from openpyxl.utils.protection import hash_password

    attrs = 'lockStructure="1"' + (f' workbookPassword="{hash_password(senha)}"' if senha else "")
    tmp = tempfile.mktemp(suffix=".xlsx")
    with zipfile.ZipFile(caminho) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                xml = data.decode("utf-8")
                xml = re.sub(r"<workbookProtection[^>]*/>", "", xml)
                xml = re.sub(r"(<workbook\b[^>]*>)", r"\1", xml, count=1)
                # workbookProtection vem logo antes de <bookViews>
                xml = xml.replace("<bookViews>", f"<workbookProtection {attrs}/><bookViews>", 1)
                xml = re.sub(r'(<sheet name="Listas"[^>]*?)state="hidden"', r'\1state="veryHidden"', xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, caminho)


if __name__ == "__main__":
    # 1) python gerar_planilha.py            -> gera os arquivos
    # 2) recalcule cada .xlsx (ex.: scripts/recalc.py do LibreOffice)
    # 3) python gerar_planilha.py --finalizar -> recoloca a trava da estrutura
    # A senha vem de PLANILHA_SENHA (opcional).
    senha = os.environ.get("PLANILHA_SENHA") or None
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    base = args[0] if args else "."
    arquivos = (f"{base}/Organizador_Contas_Atrasadas_EXEMPLO.xlsx",
                f"{base}/Organizador_Contas_Atrasadas_EM_BRANCO.xlsx")
    if "--finalizar" in sys.argv:
        for a in arquivos:
            finalizar(a, senha)
    else:
        construir(True, arquivos[0])
        construir(False, arquivos[1])
    print("ok")
