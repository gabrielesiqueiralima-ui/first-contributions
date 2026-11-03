"""Gera o PDF "Como acessar sua planilha" para a área de membros da Kiwify.

Uso: python gerar_pdf_acesso.py [saida.pdf]
Opcional: CONTATO="seu@email.com ou @instagram" para aparecer na seção de dúvidas.
"""
import os
import sys

from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

FD = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("Sans", FD + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("Sans-Bold", FD + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFontFamily("Sans", normal="Sans", bold="Sans-Bold")

VERDE = HexColor("#2E6B5E")
VERDE_CLARO = HexColor("#DCE6DC")
BEGE = HexColor("#F3E9D6")
TEXTO = HexColor("#333333")
CINZA = HexColor("#6B6B6B")

st = {
    "titulo": ParagraphStyle("t", fontName="Sans-Bold", fontSize=24, leading=30, textColor=TEXTO, alignment=TA_CENTER),
    "sub": ParagraphStyle("s", fontName="Sans", fontSize=11.5, leading=17, textColor=CINZA, alignment=TA_CENTER),
    "h": ParagraphStyle("h", fontName="Sans-Bold", fontSize=14, leading=19, textColor=VERDE, spaceBefore=6, spaceAfter=4),
    "p": ParagraphStyle("p", fontName="Sans", fontSize=10.5, leading=16, textColor=TEXTO),
    "li": ParagraphStyle("li", fontName="Sans", fontSize=10.5, leading=16, textColor=TEXTO, leftIndent=14, bulletIndent=2),
    "nota": ParagraphStyle("n", fontName="Sans", fontSize=9.5, leading=14, textColor=CINZA),
}


def caixa(conteudo, cor=BEGE):
    t = Table([[conteudo]], colWidths=[170 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), cor),
        ("LEFTPADDING", (0, 0), (-1, -1), 12), ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    return t


def passo(n, titulo, itens):
    num = Table([[Paragraph(f'<font color="#2E6B5E"><b>{n}</b></font>',
                            ParagraphStyle("x", fontName="Sans-Bold", fontSize=16, alignment=TA_CENTER, leading=20))]],
                colWidths=[11 * mm], rowHeights=[11 * mm])
    num.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), VERDE_CLARO), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    corpo = [Paragraph(f"<b>{titulo}</b>", ParagraphStyle("pt", parent=st["p"], fontSize=12, leading=17))]
    corpo += [Paragraph(i, st["li"], bulletText="•") for i in itens]
    t = Table([[num, corpo]], colWidths=[15 * mm, 155 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    return KeepTogether(t)


def gerar(saida):
    contato = os.environ.get("CONTATO", "").strip()
    doc = SimpleDocTemplate(saida, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm,
                            topMargin=18 * mm, bottomMargin=16 * mm,
                            title="Como acessar sua planilha", author="Organizador de Contas Atrasadas")
    s = []
    s += [Paragraph("Organizador de Contas Atrasadas", st["titulo"]), Spacer(1, 4),
          Paragraph("Como acessar e usar a sua planilha", st["sub"]), Spacer(1, 14)]
    s.append(caixa(Paragraph(
        "Obrigado pela compra! Em poucos minutos você vai saber <b>quanto deve</b>, "
        "<b>qual conta pagar primeiro</b> e <b>em quantos meses fica livre das dívidas</b>.", st["p"]),
        VERDE_CLARO))
    s.append(Spacer(1, 14))

    s.append(Paragraph("Primeiro: faça a sua cópia", st["h"]))
    s.append(passo(1, "Clique no link da planilha", [
        "O link está aqui na área de membros, logo abaixo deste PDF.",
        "Entre com a sua conta Google (Gmail). Se não tiver, crie uma de graça.",
        "Vai aparecer a tela <b>“Copiar documento”</b>: clique em <b>Fazer uma cópia</b>.",
        "Pronto! A planilha agora é sua e fica salva no seu Google Drive, "
        "com o nome “Cópia de Organizador de Contas Atrasadas”.",
    ]))
    s.append(passo(2, "No celular", [
        "Instale o app <b>Planilhas Google</b> (Android ou iPhone).",
        "Faça a cópia pelo navegador (passo 1) e depois abra pelo app: ela aparece na lista de planilhas.",
    ]))
    s.append(Spacer(1, 6))

    s.append(Paragraph("Depois: preencha em 4 passos", st["h"]))
    s.append(Paragraph("Você só digita nas <b>células brancas/creme</b>. As células cinzas calculam sozinhas. "
                       "Se aparecer um aviso ao tentar editar uma célula, é porque ela tem fórmula: "
                       "clique em <b>Cancelar</b>.", st["p"]))
    s.append(Spacer(1, 8))
    s.append(passo(3, "Aba Painel", [
        "Preencha seu <b>nome</b>, sua <b>renda mensal</b> e <b>quanto pode pagar por mês</b> nas dívidas.",
        "Esse último valor é o que define em quantos meses você fica livre.",
    ]))
    s.append(passo(4, "Aba Contas", [
        "Uma linha para cada conta atrasada: nome, tipo, vencimento, valor original e juros/multas.",
        "Escolha a <b>prioridade</b>: Alta para o essencial (luz, água, aluguel) e para juros muito altos "
        "(cartão, cheque especial).",
        "Quando a conta é quitada, o nome fica <s>riscado</s> sozinho.",
    ]))
    s.append(passo(5, "Aba Pagamentos", [
        "Pagou alguma coisa? Anote a data, escolha a conta na lista, o valor e a forma de pagamento.",
        "Pode lançar vários pagamentos para a mesma conta. O valor é descontado automaticamente.",
    ]))
    s.append(passo(6, "Aba Plano de Quitação", [
        "Mostra a ordem de pagamento e o mês em que cada conta fica quitada.",
        "Escolha o mês de início e a estratégia: <b>Prioridade e atraso</b> (recomendado) "
        "ou <b>Menor dívida primeiro</b>.",
        "O plano se refaz sozinho sempre que você lança um pagamento ou muda o valor por mês.",
    ]))

    s.append(Paragraph("Dicas para sair das dívidas mais rápido", st["h"]))
    for d in [
        "Comece pelas contas de prioridade <b>Alta</b> da lista <b>Pague Primeiro</b>, no Painel.",
        "Antes de pagar, <b>ligue e negocie</b>: muitos credores dão desconto para pagamento à vista ou parcelado. "
        "Anote o acordo nas colunas Negociação e Observações.",
        "Fechou um acordo com valor novo? Coloque o valor do acordo em <b>Valor original</b> e zere os juros.",
        "Atualize a planilha toda vez que pagar algo: ver o progresso subir ajuda a manter o foco.",
    ]:
        s.append(Paragraph(d, st["li"], bulletText="•"))
    s.append(Spacer(1, 10))

    s.append(Paragraph("Usa Excel?", st["h"]))
    s.append(Paragraph("Baixe o arquivo <b>.xlsx</b> disponível na área de membros e abra no Excel. "
                       "Funciona igual: preencha só as células brancas/creme.", st["p"]))
    s.append(Spacer(1, 10))

    duvida = (f"Fale com a gente: <b>{contato}</b>" if contato
              else "Responda o e-mail de confirmação da sua compra que a gente te ajuda.")
    s.append(caixa(Paragraph(f"<b>Ficou com alguma dúvida?</b> {duvida}", st["p"])))
    s.append(Spacer(1, 8))
    s.append(Paragraph("A planilha organiza e calcula com base nos valores que você informa. "
                       "Ela não substitui orientação financeira profissional e não prevê juros futuros.", st["nota"]))
    doc.build(s)


if __name__ == "__main__":
    gerar(sys.argv[1] if len(sys.argv) > 1 else "Como_acessar_sua_planilha.pdf")
    print("ok")
