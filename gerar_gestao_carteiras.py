#!/usr/bin/env python3
"""
Planilha de Gestão de Carteiras de Clientes (assessor de investimentos).

Abas:
  1. Clientes          – cadastro mestre (Penumper, Nome, CPF).
  2. Aplicacoes        – base de aplicações; Nome e CPF vêm por PROCV do
                         cadastro usando o Penumper como chave.
  3. Carteira do Cliente – escolhe um cliente e mostra a posição atual:
                         patrimônio, alocação por classe e lista de aplicações.
  4. Carteira Sugerida – montagem manual da carteira recomendada, com
                         comparativo (atual x sugerida) por classe.
  5. Instruções        – como usar.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import DoughnutChart, BarChart, PieChart, LineChart, AreaChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────
# PARÂMETROS
# ─────────────────────────────────────────
LIN_CLI = 1000          # linhas de dados na aba Clientes  (3 .. 3+LIN_CLI-1)
LIN_APL = 1000          # linhas de dados na aba Aplicacoes
LIN_HOLD = 60           # linhas da lista de aplicações por cliente
LIN_PROP = 25           # linhas para montar a carteira proposta
ANOS_PROJ = 40          # horizonte do gráfico de projeção patrimonial
ANOS_TBL = 8            # colunas da tabela periódica (anos)

CLI_FIM = 2 + LIN_CLI            # última linha de dados em Clientes
APL_FIM = 2 + LIN_APL            # última linha de dados em Aplicacoes

# Classes pré-cadastradas: (rótulo, filtro "contém" aplicado à coluna Produto)
CLASSES = [
    ("Fundos",                 "FUNDOS"),
    ("Crédito Privado",        "CRED"),
    ("Letras (LCI/LCA)",       "LETRAS"),
    ("CDB",                    "CDB"),
    ("LCA (emissor direto)",   "LCA"),
    ("Tesouraria / Outros",    "TESOURARIA"),
    ("",                       ""),
    ("",                       ""),
]
N_CLASSE = len(CLASSES)

# ─────────────────────────────────────────
# ESTILO
# ─────────────────────────────────────────
NAVY    = '1F4E79'
AZUL    = '2E75B6'
CINZA_H = 'F0F2F5'
ZEBRA   = 'F7F9FC'
AMARELO = 'FFF7D6'
VERDE   = 'E2EFDA'
BRANCO  = 'FFFFFF'

FMT_BRL = 'R$ #,##0.00;[Red]-R$ #,##0.00'
FMT_PCT = '0.0%'
FMT_DAT = 'DD/MM/YYYY'


def cor(hexv):
    return PatternFill("solid", fgColor=hexv)


def borda():
    s = Side(style='thin', color='D9DEE5')
    return Border(left=s, right=s, top=s, bottom=s)


def titulo(ws, texto, col_fim):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_fim)
    c = ws.cell(row=1, column=1, value=texto)
    c.fill = cor(NAVY)
    c.font = Font(bold=True, color=BRANCO, size=14)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30


def secao(ws, row, texto, col_ini, col_fim):
    ws.merge_cells(start_row=row, start_column=col_ini,
                   end_row=row, end_column=col_fim)
    c = ws.cell(row=row, column=col_ini, value=texto)
    c.fill = cor(NAVY)
    c.font = Font(bold=True, color=BRANCO, size=10)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row].height = 20


def cab(ws, row, col, texto):
    c = ws.cell(row=row, column=col, value=texto)
    c.fill = cor(AZUL)
    c.font = Font(bold=True, color=BRANCO, size=9)
    c.alignment = Alignment(horizontal='center', vertical='center',
                            wrap_text=True)
    c.border = borda()
    return c


def rotulo(ws, row, col, texto):
    c = ws.cell(row=row, column=col, value=texto)
    c.font = Font(bold=True, size=10, color=NAVY)
    c.alignment = Alignment(horizontal='left', vertical='center')
    return c


def campo(ws, row, col, fmt=None, destaque=True):
    c = ws.cell(row=row, column=col)
    c.fill = cor(AMARELO if destaque else BRANCO)
    c.border = borda()
    c.font = Font(size=10)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    if fmt:
        c.number_format = fmt
    return c


# ═══════════════════════════════════════════════════════════════════════════
def criar():
    wb = openpyxl.Workbook()

    # ─────────────────────────────────────────────────────────────────────
    # ABA 1 — CLIENTES
    # ─────────────────────────────────────────────────────────────────────
    cli = wb.active
    cli.title = "Clientes"
    cli.sheet_view.showGridLines = False
    titulo(cli, "CADASTRO DE CLIENTES", 3)

    for col, txt, larg in ((1, "Penumper", 18), (2, "Nome", 46),
                           (3, "CPF", 20)):
        cab(cli, 2, col, txt)
        cli.column_dimensions[get_column_letter(col)].width = larg

    for r in range(3, CLI_FIM + 1):
        zebra = ZEBRA if r % 2 else BRANCO
        for col in (1, 2, 3):
            c = cli.cell(row=r, column=col)
            c.fill = cor(zebra)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(horizontal='left', vertical='center',
                                    indent=1)
        cli.cell(row=r, column=1).number_format = '@'   # Penumper = texto
        cli.cell(row=r, column=3).number_format = '@'   # CPF = texto

    # marca Penumper / CPF duplicados em vermelho
    vermelho = Font(color='C0392B', bold=True)
    fill_verm = cor('F9D7D4')
    cli.conditional_formatting.add(
        f"A3:A{CLI_FIM}",
        FormulaRule(formula=[f'AND(A3<>"",COUNTIF($A$3:$A${CLI_FIM},A3)>1)'],
                    font=vermelho, fill=fill_verm))
    cli.conditional_formatting.add(
        f"C3:C{CLI_FIM}",
        FormulaRule(formula=[f'AND(C3<>"",COUNTIF($C$3:$C${CLI_FIM},C3)>1)'],
                    font=vermelho, fill=fill_verm))

    cli.freeze_panes = "A3"
    tb = Table(displayName="tblClientes", ref=f"A2:C{CLI_FIM}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleLight9",
                                       showRowStripes=True)
    cli.add_table(tb)

    # ─────────────────────────────────────────────────────────────────────
    # ABA 2 — APLICACOES
    # ─────────────────────────────────────────────────────────────────────
    apl = wb.create_sheet("Aplicacoes")
    apl.sheet_view.showGridLines = False
    titulo(apl, "BASE DE APLICAÇÕES — POSIÇÃO ATUAL", 10)

    colunas = [
        ("Penumper", 14, '@'),
        ("Mês\n(AAAAMM)", 11, '0'),
        ("Produto", 22, None),
        ("Subproduto", 34, None),
        ("Vencimento", 13, FMT_DAT),
        ("Saldo Ponta", 15, FMT_BRL),
        ("Idade", 9, '0'),
        ("Nome\n(automático)", 34, None),
        ("CPF\n(automático)", 18, '@'),
        ("rk", 6, '0'),
    ]
    for i, (txt, larg, _) in enumerate(colunas, 1):
        cab(apl, 2, i, txt)
        apl.column_dimensions[get_column_letter(i)].width = larg
    apl.row_dimensions[2].height = 30

    cli_rng = f"Clientes!$A$3:$C${CLI_FIM}"
    sel_pen = "'Carteira do Cliente'!$B$4"     # Penumper selecionado

    for r in range(3, APL_FIM + 1):
        zebra = ZEBRA if r % 2 else BRANCO
        for col in range(1, 11):
            c = apl.cell(row=r, column=col)
            c.fill = cor(zebra)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(horizontal='left', vertical='center',
                                    indent=1)
            fmt = colunas[col - 1][2]
            if fmt:
                c.number_format = fmt
        # colunas calculadas
        apl.cell(row=r, column=8).value = (
            f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},{cli_rng},2,0),'
            f'"NÃO CADASTRADO"))')
        apl.cell(row=r, column=9).value = (
            f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},{cli_rng},3,0),'
            f'"NÃO CADASTRADO"))')
        # rk: ranking sequencial das linhas do cliente selecionado
        apl.cell(row=r, column=10).value = (
            f'=IF($A{r}="","",IF($A{r}={sel_pen},'
            f'COUNTIF($A$3:$A{r},{sel_pen}),""))')
        for col in (8, 9):
            apl.cell(row=r, column=col).font = Font(size=10, italic=True,
                                                    color='5B6470')

    # destaca "NÃO CADASTRADO"
    apl.conditional_formatting.add(
        f"H3:I{APL_FIM}",
        FormulaRule(formula=['H3="NÃO CADASTRADO"'],
                    font=vermelho, fill=fill_verm))

    apl.column_dimensions['J'].hidden = True       # coluna auxiliar rk
    apl.freeze_panes = "A3"
    tb = Table(displayName="tblAplicacoes", ref=f"A2:J{APL_FIM}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleLight9",
                                       showRowStripes=True)
    apl.add_table(tb)

    # referências usadas pelas demais abas
    A_PEN = f"Aplicacoes!$A$3:$A${APL_FIM}"
    A_PRD = f"Aplicacoes!$C$3:$C${APL_FIM}"
    A_SUB = f"Aplicacoes!$D$3:$D${APL_FIM}"
    A_VEN = f"Aplicacoes!$E$3:$E${APL_FIM}"
    A_SLD = f"Aplicacoes!$F$3:$F${APL_FIM}"
    A_IDA = f"Aplicacoes!$G$3:$G${APL_FIM}"
    A_RK  = f"Aplicacoes!$J$3:$J${APL_FIM}"

    # ─────────────────────────────────────────────────────────────────────
    # ABA 3 — CARTEIRA DO CLIENTE
    # ─────────────────────────────────────────────────────────────────────
    car = wb.create_sheet("Carteira do Cliente")
    car.sheet_view.showGridLines = False
    titulo(car, "CARTEIRA DO CLIENTE — POSIÇÃO ATUAL", 6)
    for col, larg in ((1, 30), (2, 18), (3, 16), (4, 13), (5, 12), (6, 10)):
        car.column_dimensions[get_column_letter(col)].width = larg

    # bloco de identificação
    rotulo(car, 3, 1, "Cliente:")
    car.merge_cells("B3:D3")
    sel = campo(car, 3, 2)
    sel.font = Font(bold=True, size=11, color=NAVY)
    rotulo(car, 4, 1, "Penumper:")
    rotulo(car, 5, 1, "CPF:")
    rotulo(car, 6, 1, "Patrimônio total:")
    rotulo(car, 7, 1, "Nº de aplicações:")

    cli_nome = f"Clientes!$B$3:$B${CLI_FIM}"
    cli_pen  = f"Clientes!$A$3:$A${CLI_FIM}"
    cli_cpf  = f"Clientes!$C$3:$C${CLI_FIM}"
    for r, formula, fmt in (
            (4, f'=IFERROR(INDEX({cli_pen},MATCH($B$3,{cli_nome},0)),"")', '@'),
            (5, f'=IFERROR(INDEX({cli_cpf},MATCH($B$3,{cli_nome},0)),"")', '@'),
            (6, f'=SUMIFS({A_SLD},{A_PEN},$B$4)', FMT_BRL),
            (7, f'=COUNTIF({A_PEN},$B$4)', '0')):
        c = campo(car, r, 2, fmt, destaque=False)
        c.value = formula
        c.font = Font(bold=True, size=10, color=NAVY)

    # dropdown de clientes
    dv = DataValidation(type="list", formula1=f"={cli_nome}",
                        allow_blank=True,
                        prompt="Escolha o cliente na lista")
    car.add_data_validation(dv)
    dv.add(car["B3"])

    # ---- alocação por classe -------------------------------------------
    LIN_CL0 = 11
    secao(car, 9, "ALOCAÇÃO POR PRODUTO / CLASSE", 1, 6)
    for col, txt in ((1, "Classe"), (2, "Filtro (contém)"),
                     (3, "Saldo (R$)"), (4, "% da carteira")):
        cab(car, 10, col, txt)
    for i, (lab, filt) in enumerate(CLASSES):
        r = LIN_CL0 + i
        zebra = ZEBRA if i % 2 else BRANCO
        a = car.cell(row=r, column=1, value=lab)
        b = car.cell(row=r, column=2, value=filt)
        for col in (1, 2):
            cc = car.cell(row=r, column=col)
            cc.fill = cor(AMARELO)
            cc.border = borda()
            cc.font = Font(size=10)
            cc.alignment = Alignment(horizontal='left', vertical='center',
                                     indent=1)
        cc = car.cell(row=r, column=3,
                      value=(f'=IF($B{r}="",0,SUMIFS({A_SLD},{A_PEN},$B$4,'
                             f'{A_PRD},"*"&$B{r}&"*"))'))
        cc.number_format = FMT_BRL
        cd = car.cell(row=r, column=4,
                      value=f'=IFERROR($C{r}/$B$6,0)')
        cd.number_format = FMT_PCT
        for col in (3, 4):
            cc = car.cell(row=r, column=col)
            cc.fill = cor(zebra)
            cc.border = borda()
            cc.font = Font(size=10)
            cc.alignment = Alignment(horizontal='right', vertical='center')
    r_tot = LIN_CL0 + N_CLASSE
    r_ncl = r_tot + 1
    for r, lab, f3, f4 in (
        (r_tot, "TOTAL CLASSIFICADO",
         f'=SUM(C{LIN_CL0}:C{r_tot-1})', f'=SUM(D{LIN_CL0}:D{r_tot-1})'),
        (r_ncl, "Não classificado",
         f'=$B$6-$C${r_tot}', f'=IFERROR($C{r_ncl}/$B$6,0)')):
        ca = car.cell(row=r, column=1, value=lab)
        ca.font = Font(bold=True, size=10, color=NAVY)
        ca.alignment = Alignment(horizontal='left', vertical='center',
                                 indent=1)
        c3 = car.cell(row=r, column=3, value=f3)
        c3.number_format = FMT_BRL
        c4 = car.cell(row=r, column=4, value=f4)
        c4.number_format = FMT_PCT
        for col in (1, 2, 3, 4):
            cc = car.cell(row=r, column=col)
            cc.fill = cor(CINZA_H)
            cc.border = borda()
            if col in (3, 4):
                cc.font = Font(bold=True, size=10, color=NAVY)
                cc.alignment = Alignment(horizontal='right',
                                         vertical='center')

    # gráfico de rosca da alocação
    rosca = DoughnutChart()
    rosca.title = "Alocação por classe"
    rosca.height, rosca.width = 8, 11
    dados = Reference(car, min_col=3, min_row=10, max_row=r_tot - 1)
    cats = Reference(car, min_col=1, min_row=LIN_CL0, max_row=r_tot - 1)
    rosca.add_data(dados, titles_from_data=True)
    rosca.set_categories(cats)
    rosca.dataLabels = DataLabelList()
    rosca.dataLabels.showPercent = True
    car.add_chart(rosca, "F3")

    # ---- lista de aplicações do cliente --------------------------------
    LIN_H0 = r_ncl + 3
    secao(car, LIN_H0 - 1, "APLICAÇÕES DO CLIENTE", 1, 6)
    for col, txt in ((1, "Produto"), (2, "Subproduto"), (3, "Vencimento"),
                     (4, "Saldo (R$)"), (5, "% carteira"), (6, "Idade")):
        cab(car, LIN_H0, col, txt)
    for k in range(1, LIN_HOLD + 1):
        r = LIN_H0 + k
        zebra = ZEBRA if k % 2 else BRANCO
        defs = [
            (1, f'=IFERROR(INDEX({A_PRD},MATCH({k},{A_RK},0)),"")', None),
            (2, f'=IFERROR(INDEX({A_SUB},MATCH({k},{A_RK},0)),"")', None),
            (3, f'=IFERROR(INDEX({A_VEN},MATCH({k},{A_RK},0)),"")', FMT_DAT),
            (4, f'=IFERROR(INDEX({A_SLD},MATCH({k},{A_RK},0)),"")', FMT_BRL),
            (5, f'=IFERROR(INDEX({A_SLD},MATCH({k},{A_RK},0))/$B$6,"")',
             FMT_PCT),
            (6, f'=IFERROR(INDEX({A_IDA},MATCH({k},{A_RK},0)),"")', '0'),
        ]
        for col, formula, fmt in defs:
            c = car.cell(row=r, column=col, value=formula)
            c.fill = cor(zebra)
            c.border = borda()
            c.font = Font(size=10)
            align = 'right' if col in (3, 4, 5, 6) else 'left'
            c.alignment = Alignment(horizontal=align, vertical='center',
                                    indent=1 if align == 'left' else 0)
            if fmt:
                c.number_format = fmt

    car.freeze_panes = "A2"

    # ─────────────────────────────────────────────────────────────────────
    # ABA 4 — CARTEIRA SUGERIDA  (proposta com classes, liquidez, racional)
    # ─────────────────────────────────────────────────────────────────────
    sug = wb.create_sheet("Carteira Sugerida")
    sug.sheet_view.showGridLines = False
    titulo(sug, "PROPOSTA DE CARTEIRA", 14)
    larguras_sug = [22, 28, 12, 12, 16, 10, 2,
                    12, 12, 12, 12, 12, 12, 12]
    for i, larg in enumerate(larguras_sug, 1):
        sug.column_dimensions[get_column_letter(i)].width = larg

    # ---- cabeçalho ------------------------------------------------------
    rotulo(sug, 3, 1, "Cliente:")
    sug.merge_cells("B3:E3")
    c = campo(sug, 3, 2, destaque=False)
    c.value = "='Carteira do Cliente'!B3"
    c.font = Font(bold=True, size=12, color=NAVY)
    rotulo(sug, 4, 1, "Patrimônio Atual:")
    sug.merge_cells("B4:E4")
    c = campo(sug, 4, 2, FMT_BRL, destaque=False)
    c.value = "='Carteira do Cliente'!B6"
    c.font = Font(bold=True, size=10, color=NAVY)

    # banner direito: "Patrimônio Proposto"
    VERMELHO_BANNER = 'C0392B'
    sug.merge_cells("H3:N3")
    c = sug.cell(row=3, column=8, value="PATRIMÔNIO PROPOSTO")
    c.fill = cor(VERMELHO_BANNER)
    c.font = Font(bold=True, color=BRANCO, size=11)
    c.alignment = Alignment(horizontal='center', vertical='center')
    sug.merge_cells("H4:N4")
    c = sug.cell(row=4, column=8, value="=SUM(E8:E32)")
    c.fill = cor(NAVY)
    c.font = Font(bold=True, color=BRANCO, size=16)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.number_format = FMT_BRL
    sug.row_dimensions[3].height = 22
    sug.row_dimensions[4].height = 30

    # ---- tabela carteira recomendada -----------------------------------
    secao(sug, 6, "CARTEIRA RECOMENDADA", 1, 6)
    heads_prop = ["Classe", "Ativo sugerido", "Carência\n(dias corridos)",
                  "Liquidez\n(D + dias)", "Valor (R$)", "%"]
    for col, txt in enumerate(heads_prop, 1):
        cab(sug, 7, col, txt)
    sug.row_dimensions[7].height = 32

    for i in range(LIN_PROP):
        r = 8 + i
        for col in range(1, 6):
            c = sug.cell(row=r, column=col)
            c.fill = cor(AMARELO)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(
                horizontal='right' if col >= 3 else 'left',
                vertical='center', indent=0 if col >= 3 else 1)
        sug.cell(row=r, column=3).number_format = '0'
        sug.cell(row=r, column=4).number_format = '0'
        sug.cell(row=r, column=5).number_format = FMT_BRL
        cf = sug.cell(row=r, column=6,
                      value=f'=IFERROR($E{r}/SUM($E$8:$E$32),"")')
        cf.fill = cor(BRANCO)
        cf.border = borda()
        cf.font = Font(size=10)
        cf.number_format = FMT_PCT
        cf.alignment = Alignment(horizontal='right', vertical='center')

    r_prop_fim = 8 + LIN_PROP - 1        # 32
    r_prop_tot = r_prop_fim + 1          # 33
    for col, val, fmt in (
        (1, "TOTAL", None),
        (5, f"=SUM(E8:E{r_prop_fim})", FMT_BRL),
        (6, f"=SUM(F8:F{r_prop_fim})", FMT_PCT),
    ):
        c = sug.cell(row=r_prop_tot, column=col, value=val)
        c.fill = cor(CINZA_H)
        c.border = borda()
        c.font = Font(bold=True, size=10, color=NAVY)
        c.alignment = Alignment(
            horizontal='right' if col >= 3 else 'left',
            vertical='center', indent=1)
        if fmt:
            c.number_format = fmt
    for col in (2, 3, 4):
        c = sug.cell(row=r_prop_tot, column=col)
        c.fill = cor(CINZA_H)
        c.border = borda()

    # dropdown de classe na proposta
    dv2 = DataValidation(
        type="list",
        formula1=f"='Carteira do Cliente'!$A${LIN_CL0}:$A${LIN_CL0+N_CLASSE-1}",
        allow_blank=True)
    sug.add_data_validation(dv2)
    dv2.add(f"A8:A{r_prop_fim}")

    # ---- helpers ocultos para os gráficos -------------------------------
    # cols Q, R: agregação por classe (pizza)
    # cols S, T: agregação por liquidez (barras)
    for col_letter in ('Q', 'R', 'S', 'T'):
        sug.column_dimensions[col_letter].hidden = True
    sug.cell(row=7, column=17, value="Classe").font = Font(bold=True, size=9)
    sug.cell(row=7, column=18, value="Valor").font = Font(bold=True, size=9)
    for i in range(N_CLASSE):
        r = 8 + i
        cr = LIN_CL0 + i
        sug.cell(row=r, column=17, value=f"='Carteira do Cliente'!A{cr}")
        sug.cell(row=r, column=18,
                 value=(f'=IF($Q{r}="",0,SUMIFS($E$8:$E${r_prop_fim},'
                        f'$A$8:$A${r_prop_fim},$Q{r}))'))
    sug.cell(row=7, column=19, value="D+").font = Font(bold=True, size=9)
    sug.cell(row=7, column=20, value="Valor").font = Font(bold=True, size=9)
    buckets = [0, 1, 10, 30, 45, 90, 180, 270, 365, 720, 1800]
    for i, b in enumerate(buckets):
        r = 8 + i
        sug.cell(row=r, column=19, value=b)
        sug.cell(row=r, column=20,
                 value=(f'=SUMIFS($E$8:$E${r_prop_fim},'
                        f'$D$8:$D${r_prop_fim},$S{r})'))

    # gráfico de pizza – composição por classe
    pizza = PieChart()
    pizza.title = "Composição por Classe de Ativos"
    pizza.height, pizza.width = 7, 9.5
    pizza.add_data(Reference(sug, min_col=18, min_row=8,
                             max_row=7 + N_CLASSE), titles_from_data=False)
    pizza.set_categories(Reference(sug, min_col=17, min_row=8,
                                    max_row=7 + N_CLASSE))
    pizza.dataLabels = DataLabelList()
    pizza.dataLabels.showPercent = True
    sug.add_chart(pizza, "H6")

    # gráfico de barras – distribuição de liquidez
    barra_liq = BarChart()
    barra_liq.type = 'col'
    barra_liq.title = "Distribuição de Liquidez (R$ por D+)"
    barra_liq.height, barra_liq.width = 7, 9.5
    barra_liq.y_axis.numFmt = 'R$ #,##0'
    barra_liq.x_axis.title = "Liquidez (dias)"
    barra_liq.x_axis.delete = False
    barra_liq.y_axis.delete = False
    barra_liq.add_data(Reference(sug, min_col=20, min_row=8,
                                  max_row=7 + len(buckets)),
                       titles_from_data=False)
    barra_liq.set_categories(Reference(sug, min_col=19, min_row=8,
                                        max_row=7 + len(buckets)))
    barra_liq.legend = None
    sug.add_chart(barra_liq, "H22")

    # ---- otimizações realizadas (texto livre) --------------------------
    R_OT = r_prop_tot + 4                # 37
    secao(sug, R_OT, "OTIMIZAÇÕES REALIZADAS", 1, 14)
    for r in range(R_OT + 1, R_OT + 6):
        sug.merge_cells(start_row=r, start_column=1,
                        end_row=r, end_column=14)
        c = sug.cell(row=r, column=1)
        c.fill = cor(AMARELO)
        c.border = borda()
        c.font = Font(size=10)
        c.alignment = Alignment(horizontal='left', vertical='top',
                                wrap_text=True, indent=1)
        sug.row_dimensions[r].height = 22

    # ---- racional por movimentação --------------------------------------
    R_RC = R_OT + 7                      # 44
    secao(sug, R_RC, "RACIONAL POR MOVIMENTAÇÃO", 1, 14)
    cab(sug, R_RC + 1, 1, "Ativo")
    sug.merge_cells(start_row=R_RC + 1, start_column=2,
                    end_row=R_RC + 1, end_column=14)
    cab(sug, R_RC + 1, 2, "Racional (objetivo e justificativa)")
    for i in range(10):
        r = R_RC + 2 + i
        c = sug.cell(row=r, column=1)
        c.fill = cor(AMARELO)
        c.border = borda()
        c.font = Font(size=10)
        c.alignment = Alignment(horizontal='left', vertical='center',
                                wrap_text=True, indent=1)
        sug.merge_cells(start_row=r, start_column=2,
                        end_row=r, end_column=14)
        c2 = sug.cell(row=r, column=2)
        c2.fill = cor(AMARELO)
        c2.border = borda()
        c2.font = Font(size=10)
        c2.alignment = Alignment(horizontal='left', vertical='top',
                                  wrap_text=True, indent=1)
        sug.row_dimensions[r].height = 32

    # ---- comparativo atual x sugerida ----------------------------------
    R_CMP = R_RC + 13                    # 57
    secao(sug, R_CMP, "COMPARATIVO: ATUAL x SUGERIDA", 1, 7)
    heads = ["Classe", "% Atual", "% Sugerido", "Saldo Atual (R$)",
             "Valor Sugerido (R$)", "Diferença (R$)", "Ação"]
    for col, txt in enumerate(heads, 1):
        cab(sug, R_CMP + 1, col, txt)
    cmp0 = R_CMP + 2
    for i in range(N_CLASSE):
        r = cmp0 + i
        zebra = ZEBRA if i % 2 else BRANCO
        cr = LIN_CL0 + i
        vals = [
            (1, f"='Carteira do Cliente'!A{cr}", None),
            (2, f"='Carteira do Cliente'!D{cr}", FMT_PCT),
            (3, f'=IFERROR($E{r}/SUM($E$8:$E${r_prop_fim}),0)', FMT_PCT),
            (4, f"='Carteira do Cliente'!C{cr}", FMT_BRL),
            (5, (f'=SUMIFS($E$8:$E${r_prop_fim},'
                 f'$A$8:$A${r_prop_fim},$A{r})'), FMT_BRL),
            (6, f'=$E{r}-$D{r}', FMT_BRL),
            (7, (f'=IF(AND($D{r}=0,$E{r}=0),"",'
                 f'IF(ROUND($F{r},2)>0,"Aumentar",'
                 f'IF(ROUND($F{r},2)<0,"Reduzir","Manter")))'), None),
        ]
        for col, formula, fmt in vals:
            c = sug.cell(row=r, column=col, value=formula)
            c.fill = cor(zebra)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(
                horizontal='left' if col in (1, 7) else 'right',
                vertical='center', indent=1 if col in (1, 7) else 0)
            if fmt:
                c.number_format = fmt
    r_cmp_tot = cmp0 + N_CLASSE
    for col, formula in ((1, "TOTAL"),
                         (4, f"=SUM(D{cmp0}:D{r_cmp_tot-1})"),
                         (5, f"=SUM(E{cmp0}:E{r_cmp_tot-1})"),
                         (6, f"=SUM(F{cmp0}:F{r_cmp_tot-1})")):
        c = sug.cell(row=r_cmp_tot, column=col, value=formula)
        c.fill = cor(CINZA_H)
        c.border = borda()
        c.font = Font(bold=True, size=10, color=NAVY)
        c.alignment = Alignment(horizontal='left' if col == 1 else 'right',
                                vertical='center', indent=1)
        if col in (4, 5, 6):
            c.number_format = FMT_BRL

    sug.conditional_formatting.add(
        f"G{cmp0}:G{r_cmp_tot-1}",
        FormulaRule(formula=[f'$G{cmp0}="Aumentar"'],
                    font=Font(color='1E7B34', bold=True),
                    fill=cor(VERDE)))
    sug.conditional_formatting.add(
        f"G{cmp0}:G{r_cmp_tot-1}",
        FormulaRule(formula=[f'$G{cmp0}="Reduzir"'],
                    font=vermelho, fill=fill_verm))

    sug.freeze_panes = "A6"

    # ─────────────────────────────────────────────────────────────────────
    # ABA 5 — PLANEJAMENTO  (premissas + projeção patrimonial)
    # ─────────────────────────────────────────────────────────────────────
    pln = wb.create_sheet("Planejamento")
    pln.sheet_view.showGridLines = False
    titulo(pln, "PLANEJAMENTO FINANCEIRO & APOSENTADORIA", 10)
    larguras_pln = [38, 16, 6, 38, 16, 4, 16, 16, 16, 14]
    for i, larg in enumerate(larguras_pln, 1):
        pln.column_dimensions[get_column_letter(i)].width = larg

    rotulo(pln, 3, 1, "Cliente:")
    pln.merge_cells("B3:D3")
    c = campo(pln, 3, 2, destaque=False)
    c.value = "='Carteira do Cliente'!B3"
    c.font = Font(bold=True, size=12, color=NAVY)

    # ---- INFORMAÇÕES GERAIS E PATRIMÔNIO ATUAL ---------------------------
    secao(pln, 5, "INFORMAÇÕES GERAIS E PATRIMÔNIO ATUAL", 1, 2)
    gerais = [
        ("Idade atual (anos)", 51, '0'),
        ("Idade de aposentadoria (anos)", 64, '0'),
        ("Expectativa de vida (anos)", 90, '0'),
        ("Horizonte de produção/acumulação (anos)", "=B7-B6", '0'),
        ("Patrimônio líquido disponível (R$)", 0, FMT_BRL),
        ("Aplicação mensal (Prev + Meta) (R$)", 0, FMT_BRL),
    ]
    for i, (lab, val, fmt) in enumerate(gerais):
        r = 6 + i
        rotulo(pln, r, 1, lab)
        c = campo(pln, r, 2, fmt,
                  destaque=isinstance(val, (int, float)))
        c.value = val
    # nomeia para clareza: B6 IdadeAtual, B7 IdadeApo, B8 ExpVida, B9 Horiz
    # B10 Patrim, B11 AplicMes

    # ---- INFLAÇÃO, CDI, ALÍQUOTA IR E TAXA NOMINAL -----------------------
    secao(pln, 5, "INFLAÇÃO, CDI, ALÍQUOTA IR E TAXA NOMINAL", 4, 5)
    econ = [
        ("Inflação anual projetada", 0.04, '0.00%'),
        ("Alíquota de IR (sobre rendimentos)", 0.15, '0.00%'),
        ("Taxa CDI anual", 0.1065, '0.00%'),
        ("% do CDI esperado (rentabilidade)", 0.72, '0.00%'),
        ("Taxa nominal anual bruta (calc)", "=E8*E9", '0.00%'),
        ("Taxa nominal anual líquida (calc)",
         "=E10*(1-E7)", '0.00%'),
    ]
    for i, (lab, val, fmt) in enumerate(econ):
        r = 6 + i
        rotulo(pln, r, 4, lab)
        is_input = not (isinstance(val, str) and val.startswith('='))
        c = campo(pln, r, 5, fmt, destaque=is_input)
        c.value = val
    # E6 inflação, E7 IR, E8 CDI, E9 %CDI, E10 nominal_bruta, E11 nominal_liq

    # ---- PREMISSAS DE APOSENTADORIA --------------------------------------
    secao(pln, 13, "PREMISSAS DE APOSENTADORIA", 1, 2)
    apos = [
        ("Renda desejada na aposentadoria (R$ / mês)", 0, FMT_BRL),
        ("Renda projetada do INSS (R$ / mês)", 0, FMT_BRL),
        ("Outras fontes de renda (R$ / mês)", 0, FMT_BRL),
        ("Renda a sustentar pelo patrimônio (R$ / mês)",
         "=MAX(0,B14-B15-B16)", FMT_BRL),
        ("Renda a sustentar pelo patrimônio (R$ / ano)",
         "=B17*12", FMT_BRL),
    ]
    for i, (lab, val, fmt) in enumerate(apos):
        r = 14 + i
        rotulo(pln, r, 1, lab)
        is_input = not (isinstance(val, str) and val.startswith('='))
        c = campo(pln, r, 2, fmt, destaque=is_input)
        c.value = val
    # B14 renda desejada, B15 INSS, B16 outras, B17 sustentar mensal, B18 anual

    # ---- TAXA REAL --------------------------------------------------------
    secao(pln, 13, "TAXA DE JUROS REAL", 4, 5)
    real = [
        ("Taxa real anual bruta",
         "=(1+E10)/(1+E6)-1", '0.00%'),
        ("Taxa real anual líquida",
         "=(1+E11)/(1+E6)-1", '0.00%'),
        ("Taxa real anual líquida pós-aposentadoria",
         0.03, '0.00%'),
        ("Taxa real mensal líquida pós-aposentadoria",
         "=(1+E16)^(1/12)-1", '0.0000%'),
    ]
    for i, (lab, val, fmt) in enumerate(real):
        r = 14 + i
        rotulo(pln, r, 4, lab)
        is_input = not (isinstance(val, str) and val.startswith('='))
        c = campo(pln, r, 5, fmt, destaque=is_input)
        c.value = val
    # E14 real anual bruta, E15 real anual liq pré-apo,
    # E16 real anual liq pós-apo, E17 real mensal liq pós-apo

    # ---- APOSENTADORIA (Preservando patrimônio) -------------------------
    secao(pln, 20, "APOSENTADORIA — VIVER DE RENDA (preservando)", 1, 2)
    rotulo(pln, 21, 1, "Patrimônio necessário na aposentadoria")
    c = campo(pln, 21, 2, FMT_BRL, destaque=False)
    c.value = "=IFERROR(B18/E16,0)"
    c.font = Font(bold=True, size=10, color=NAVY)
    rotulo(pln, 22, 1, "Parcela mensal necessária a aplicar")
    c = campo(pln, 22, 2, FMT_BRL, destaque=False)
    c.value = "=IFERROR(-PMT(E17,B9*12,-B10,B21),0)"
    c.font = Font(bold=True, size=10, color=NAVY)

    # ---- APOSENTADORIA (Consumo de patrimônio com sucessão) -------------
    secao(pln, 20, "APOSENTADORIA — CONSUMO COM SUCESSÃO", 4, 5)
    rotulo(pln, 21, 4, "Patrimônio de sucessão desejado (R$)")
    c = campo(pln, 21, 5, FMT_BRL)
    c.value = 0
    rotulo(pln, 22, 4, "Patrimônio necessário na aposentadoria")
    c = campo(pln, 22, 5, FMT_BRL, destaque=False)
    c.value = ("=IFERROR(PV(E16,B8-B7,-B18,-E21),0)")
    c.font = Font(bold=True, size=10, color=NAVY)
    rotulo(pln, 23, 4, "Parcela mensal necessária a aplicar")
    c = campo(pln, 23, 5, FMT_BRL, destaque=False)
    c.value = "=IFERROR(-PMT(E17,B9*12,-B10,E22),0)"
    c.font = Font(bold=True, size=10, color=NAVY)

    # ---- OBJETIVOS DE VIDA (1-6 anos) -----------------------------------
    secao(pln, 25, "OBJETIVOS DE VIDA (1-6 anos)", 1, 5)
    cab(pln, 26, 1, "Projeto")
    cab(pln, 26, 2, "Valor hoje (R$)")
    cab(pln, 26, 4, "Acumulação projetada (R$)")
    cab(pln, 26, 5, "Anos")
    objetivos = [("Projetos 1-2 anos", 2), ("Projetos 3-4 anos", 4),
                 ("Projetos 5-6 anos", 6)]
    for i, (lab, anos) in enumerate(objetivos):
        r = 27 + i
        rotulo(pln, r, 1, lab)
        c = campo(pln, r, 2, FMT_BRL)
        c.value = 0
        c2 = campo(pln, r, 4, FMT_BRL, destaque=False)
        c2.value = f"=B{r}*(1+E10)^E{r}"
        c2.font = Font(bold=True, size=10, color=NAVY)
        c3 = campo(pln, r, 5, '0', destaque=False)
        c3.value = anos
        c3.font = Font(size=10, color=NAVY)
    rotulo(pln, 30, 1, "Total de Projetos e Acumulação")
    c = campo(pln, 30, 2, FMT_BRL, destaque=False)
    c.value = "=SUM(B27:B29)"
    c.font = Font(bold=True, size=10, color=NAVY)
    c = campo(pln, 30, 4, FMT_BRL, destaque=False)
    c.value = "=SUM(D27:D29)"
    c.font = Font(bold=True, size=10, color=NAVY)

    # ---- PROJEÇÃO PATRIMONIAL (tabela + gráfico) ------------------------
    secao(pln, 33, "PROJEÇÃO PATRIMONIAL (em R$ reais de hoje)", 7, 10)
    cab(pln, 34, 7, "Ano")
    cab(pln, 34, 8, "Idade")
    cab(pln, 34, 9, "Saldo Real")
    cab(pln, 34, 10, "Viver de Renda")
    for y in range(ANOS_PROJ):
        r = 35 + y
        zebra = ZEBRA if y % 2 else BRANCO
        pln.cell(row=r, column=7, value=y)
        pln.cell(row=r, column=8, value=f"=B6+G{r}")
        # Saldo Real: acumulação até idade aposentadoria; depois consome
        if y == 0:
            pln.cell(row=r, column=9, value="=B10")
        else:
            pln.cell(row=r, column=9,
                     value=(f'=IF(H{r}<=B7,'
                            f'I{r-1}*(1+E15)+B11*12,'
                            f'MAX(0,I{r-1}*(1+E16)-B18))'))
        # Viver de Renda: linha de referência = patrimônio necessário
        pln.cell(row=r, column=10,
                 value=f"=IF(H{r}<B7,NA(),B21)")
        for col in (7, 8, 9, 10):
            cc = pln.cell(row=r, column=col)
            cc.fill = cor(zebra)
            cc.border = borda()
            cc.font = Font(size=9)
            cc.alignment = Alignment(
                horizontal='center' if col in (7, 8) else 'right',
                vertical='center')
            if col in (9, 10):
                cc.number_format = FMT_BRL

    # gráfico projeção patrimonial
    proj = AreaChart()
    proj.title = "Projeção Patrimonial — Independência Financeira & Aposentadoria"
    proj.height, proj.width = 11, 22
    proj.y_axis.title = "Patrimônio (R$)"
    proj.x_axis.title = "Idade"
    proj.y_axis.numFmt = 'R$ #,##0'
    proj.x_axis.delete = False
    proj.y_axis.delete = False
    dref = Reference(pln, min_col=9, min_row=34,
                     max_row=34 + ANOS_PROJ)
    proj.add_data(dref, titles_from_data=True)
    proj.set_categories(Reference(pln, min_col=8, min_row=35,
                                    max_row=34 + ANOS_PROJ))
    # adiciona linha "Viver de Renda" como série de linha
    linha_ref = LineChart()
    linha_ref.add_data(Reference(pln, min_col=10, min_row=34,
                                  max_row=34 + ANOS_PROJ),
                       titles_from_data=True)
    proj += linha_ref
    proj.visible_cells_only = False
    pln.add_chart(proj, "A36")

    pln.freeze_panes = "A5"

    # ─────────────────────────────────────────────────────────────────────
    # ABA 6 — TABELA PERIÓDICA DOS INVESTIMENTOS
    # ─────────────────────────────────────────────────────────────────────
    tpr = wb.create_sheet("Tabela Periódica")
    tpr.sheet_view.showGridLines = False
    titulo(tpr, "TABELA PERIÓDICA DOS INVESTIMENTOS — RETORNO ANUAL", 12)
    tpr.column_dimensions['A'].width = 22
    for i in range(ANOS_TBL + 1):
        tpr.column_dimensions[get_column_letter(2 + i)].width = 11

    indices = ["S&P 500", "Euro Stoxx 50", "Ibovespa", "IFIX",
               "IMA-B", "IRF-M", "IHFA", "CDI", "Dólar", "Inflação"]
    ano_atual_default = 2024
    cab(tpr, 3, 1, "Índice / Classe")
    for i in range(ANOS_TBL):
        ano = ano_atual_default - ANOS_TBL + 1 + i
        cab(tpr, 3, 2 + i, str(ano))
    cab(tpr, 3, 2 + ANOS_TBL, "Acumulado")
    tpr.row_dimensions[3].height = 22

    for i, nome in enumerate(indices):
        r = 4 + i
        zebra = ZEBRA if i % 2 else BRANCO
        c = tpr.cell(row=r, column=1, value=nome)
        c.fill = cor(zebra)
        c.border = borda()
        c.font = Font(bold=True, size=10, color=NAVY)
        c.alignment = Alignment(horizontal='left', vertical='center',
                                indent=1)
        for j in range(ANOS_TBL):
            cc = tpr.cell(row=r, column=2 + j)
            cc.fill = cor(AMARELO)
            cc.border = borda()
            cc.font = Font(size=10)
            cc.alignment = Alignment(horizontal='right', vertical='center')
            cc.number_format = '0.00%;[Red]-0.00%'
        # coluna Acumulado calcula automaticamente
        cc = tpr.cell(row=r, column=2 + ANOS_TBL)
        col_ini = get_column_letter(2)
        col_fim = get_column_letter(1 + ANOS_TBL)
        cc.value = (f'=IFERROR(PRODUCT(IF(ISNUMBER({col_ini}{r}:{col_fim}{r}),'
                    f'1+{col_ini}{r}:{col_fim}{r},1))-1,"")')
        cc.fill = cor(CINZA_H)
        cc.border = borda()
        cc.font = Font(bold=True, size=10, color=NAVY)
        cc.alignment = Alignment(horizontal='right', vertical='center')
        cc.number_format = '0.00%;[Red]-0.00%'

    # color-scale: melhor da coluna em verde, pior em vermelho
    from openpyxl.formatting.rule import ColorScaleRule
    for j in range(ANOS_TBL):
        col = get_column_letter(2 + j)
        rng = f"{col}4:{col}{3 + len(indices)}"
        regra = ColorScaleRule(
            start_type='min', start_color='F8696B',
            mid_type='percentile', mid_value=50, mid_color='FFEB84',
            end_type='max', end_color='63BE7B')
        tpr.conditional_formatting.add(rng, regra)

    # observação
    obs_r = 4 + len(indices) + 2
    tpr.merge_cells(start_row=obs_r, start_column=1,
                     end_row=obs_r, end_column=2 + ANOS_TBL)
    c = tpr.cell(row=obs_r, column=1,
                 value=("Preencha os retornos anuais (em %) de cada índice/classe. "
                        "A coluna Acumulado é calculada automaticamente; a cor "
                        "destaca os melhores (verde) e piores (vermelho) de cada ano."))
    c.alignment = Alignment(horizontal='left', vertical='center',
                            wrap_text=True, indent=1)
    c.font = Font(size=9, italic=True, color='5B6470')
    tpr.row_dimensions[obs_r].height = 30

    tpr.freeze_panes = "B4"

    # ─────────────────────────────────────────────────────────────────────
    # ABA 7 — INSTRUÇÕES
    # ─────────────────────────────────────────────────────────────────────
    ins = wb.create_sheet("Instruções")
    ins.sheet_view.showGridLines = False
    ins.column_dimensions['A'].width = 105
    t = ins.cell(row=1, column=1, value="COMO USAR ESTA PLANILHA")
    t.fill = cor(NAVY)
    t.font = Font(bold=True, color=BRANCO, size=13)
    t.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ins.row_dimensions[1].height = 28

    linhas = [
        "",
        "1. Aba 'Clientes': cadastre cada cliente uma única vez — Penumper, "
        "Nome e CPF. O Penumper é a chave que liga tudo; pode ter letras "
        "(ex.: 1297217A), por isso a coluna é formatada como texto.",
        "     Penumper ou CPF repetidos ficam destacados em vermelho.",
        "",
        "2. Aba 'Aplicacoes': cole aqui a posição vinda do sistema — Penumper, "
        "Mês, Produto, Subproduto, Vencimento, Saldo Ponta e Idade.",
        "     As colunas Nome e CPF se preenchem sozinhas por PROCV usando o "
        "Penumper. Se aparecer 'NÃO CADASTRADO', falta cadastrar o cliente na "
        "aba 'Clientes'.",
        "     Cada linha é um investimento (um produto/subproduto com saldo).",
        "",
        "3. Aba 'Carteira do Cliente': escolha o cliente no menu suspenso. "
        "A planilha mostra o patrimônio total, a alocação por classe (com "
        "gráfico) e a lista das aplicações daquele cliente.",
        "     A linha 'Não classificado' indica saldo cujo Produto não casou "
        "com nenhum filtro — ajuste os filtros da coluna B se isso acontecer.",
        "",
        "4. Classes: na aba 'Carteira do Cliente' a coluna 'Filtro (contém)' "
        "define como cada classe é somada. Ex.: o filtro FUNDOS soma todo "
        "Produto que contenha 'FUNDOS'. Edite rótulos e filtros à vontade.",
        "",
        "5. Aba 'Carteira Sugerida' (proposta): monte a carteira recomendada — "
        "Classe, Ativo sugerido, Carência (dias), Liquidez (dias) e Valor. "
        "O banner 'Patrimônio Proposto' soma sozinho e os dois gráficos "
        "(composição por classe e distribuição de liquidez) se atualizam.",
        "     Logo abaixo há campos livres para você descrever as Otimizações "
        "Realizadas e o Racional por movimentação (objetivo, justificativa).",
        "     O Comparativo no fim mostra, classe a classe, o % atual x o "
        "% sugerido e quanto Aumentar ou Reduzir.",
        "",
        "6. Aba 'Planejamento': preencha os campos amarelos (idades, "
        "patrimônio, aplicação mensal, renda desejada, INSS, premissas de "
        "inflação/CDI/IR/taxa real). A planilha calcula sozinha: o patrimônio "
        "necessário na aposentadoria nos cenários 'Viver de Renda' "
        "(preservando) e 'Consumo com Sucessão', a parcela mensal a aplicar "
        "em cada cenário, os objetivos de vida (projetos 1-6 anos) e a "
        "Projeção Patrimonial em R$ de hoje (gráfico de área).",
        "",
        "7. Aba 'Tabela Periódica': preencha os retornos anuais (%) de cada "
        "índice/classe nas colunas dos anos. A coluna 'Acumulado' calcula "
        "sozinha e a cor destaca, em cada ano, o melhor (verde) e o pior "
        "(vermelho) desempenho.",
        "",
        "Observação: a planilha trabalha com a posição atual (foto). Para "
        "acrescentar uma nova posição, basta colar mais linhas na aba "
        "'Aplicacoes' — as fórmulas se estendem automaticamente.",
    ]
    for i, txt in enumerate(linhas, 2):
        c = ins.cell(row=i, column=1, value=txt)
        c.alignment = Alignment(horizontal='left', vertical='center',
                                wrap_text=True)
        neg = txt[:2].strip().rstrip('.').isdigit()
        c.font = Font(size=11, bold=neg, color=NAVY if neg else '333333')
        ins.row_dimensions[i].height = 18

    # ─────────────────────────────────────────────────────────────────────
    wb.active = car
    caminho = "/home/user/Planilha/gestao_carteiras.xlsx"
    wb.save(caminho)
    print(f"Planilha gerada: {caminho}")


if __name__ == "__main__":
    criar()
