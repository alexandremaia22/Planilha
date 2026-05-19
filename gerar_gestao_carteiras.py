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
from openpyxl.chart import DoughnutChart, BarChart, Reference
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
LIN_SUG = 25            # linhas para montar a carteira sugerida

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
    # ABA 4 — CARTEIRA SUGERIDA
    # ─────────────────────────────────────────────────────────────────────
    sug = wb.create_sheet("Carteira Sugerida")
    sug.sheet_view.showGridLines = False
    titulo(sug, "CARTEIRA SUGERIDA", 7)
    larguras = [26, 16, 13, 16, 30, 16, 13]
    for i, larg in enumerate(larguras, 1):
        sug.column_dimensions[get_column_letter(i)].width = larg

    rotulo(sug, 3, 1, "Cliente:")
    sug.merge_cells("B3:C3")
    c = campo(sug, 3, 2, destaque=False)
    c.value = "='Carteira do Cliente'!B3"
    c.font = Font(bold=True, size=11, color=NAVY)
    rotulo(sug, 4, 1, "Patrimônio atual:")
    c = campo(sug, 4, 2, FMT_BRL, destaque=False)
    c.value = "='Carteira do Cliente'!B6"
    c.font = Font(bold=True, size=10, color=NAVY)
    obs = sug.cell(row=3, column=5,
                   value="Selecione o cliente na aba 'Carteira do Cliente'.")
    obs.font = Font(size=9, italic=True, color='5B6470')

    # ---- carteira atual por classe (esquerda) --------------------------
    secao(sug, 6, "CARTEIRA ATUAL POR CLASSE", 1, 3)
    for col, txt in ((1, "Classe"), (2, "Saldo Atual (R$)"),
                     (3, "% Atual")):
        cab(sug, 7, col, txt)
    for i in range(N_CLASSE):
        r = 8 + i
        zebra = ZEBRA if i % 2 else BRANCO
        cr = LIN_CL0 + i
        vals = [
            (1, f"='Carteira do Cliente'!A{cr}", None),
            (2, f"='Carteira do Cliente'!C{cr}", FMT_BRL),
            (3, f"='Carteira do Cliente'!D{cr}", FMT_PCT),
        ]
        for col, formula, fmt in vals:
            c = sug.cell(row=r, column=col, value=formula)
            c.fill = cor(zebra)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(
                horizontal='left' if col == 1 else 'right',
                vertical='center', indent=1 if col == 1 else 0)
            if fmt:
                c.number_format = fmt
    r_at = 8 + N_CLASSE
    for col, formula in ((1, "TOTAL"),
                         (2, f"=SUM(B8:B{r_at-1})"),
                         (3, f"=SUM(C8:C{r_at-1})")):
        c = sug.cell(row=r_at, column=col, value=formula)
        c.fill = cor(CINZA_H)
        c.border = borda()
        c.font = Font(bold=True, size=10, color=NAVY)
        c.alignment = Alignment(horizontal='left' if col == 1 else 'right',
                                vertical='center', indent=1)
        if col == 2:
            c.number_format = FMT_BRL
        if col == 3:
            c.number_format = FMT_PCT

    # ---- carteira sugerida — preenchimento manual (direita) ------------
    secao(sug, 6, "CARTEIRA SUGERIDA — PREENCHA", 5, 7)
    for col, txt in ((5, "Classe"), (6, "Investimento sugerido"),
                     (7, "Valor (R$)")):
        cab(sug, 7, col, txt)
    for i in range(LIN_SUG):
        r = 8 + i
        zebra = ZEBRA if i % 2 else BRANCO
        for col in (5, 6, 7):
            c = sug.cell(row=r, column=col)
            c.fill = cor(AMARELO)
            c.border = borda()
            c.font = Font(size=10)
            c.alignment = Alignment(
                horizontal='right' if col == 7 else 'left',
                vertical='center', indent=0 if col == 7 else 1)
        sug.cell(row=r, column=7).number_format = FMT_BRL
    r_sug_fim = 8 + LIN_SUG - 1
    r_sug_tot = r_sug_fim + 1
    c = sug.cell(row=r_sug_tot, column=6, value="TOTAL SUGERIDO")
    c.font = Font(bold=True, size=10, color=NAVY)
    c.alignment = Alignment(horizontal='right', vertical='center')
    c = sug.cell(row=r_sug_tot, column=7, value=f"=SUM(G8:G{r_sug_fim})")
    c.fill = cor(CINZA_H)
    c.border = borda()
    c.font = Font(bold=True, size=10, color=NAVY)
    c.number_format = FMT_BRL
    c.alignment = Alignment(horizontal='right', vertical='center')

    # dropdown de classe na carteira sugerida
    dv2 = DataValidation(
        type="list",
        formula1=f"='Carteira do Cliente'!$A${LIN_CL0}:$A${LIN_CL0+N_CLASSE-1}",
        allow_blank=True)
    sug.add_data_validation(dv2)
    dv2.add(f"E8:E{r_sug_fim}")

    # ---- comparativo por classe ----------------------------------------
    # abaixo da seção de preenchimento (evita sobreposição de linhas)
    R_CMP = max(r_at, r_sug_tot) + 3
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
            (3, f'=IFERROR($E{r}/$G${r_sug_tot},0)', FMT_PCT),
            (4, f"='Carteira do Cliente'!C{cr}", FMT_BRL),
            (5, (f'=SUMIFS($G$8:$G${r_sug_fim},$E$8:$E${r_sug_fim},'
                 f'$A{r})'), FMT_BRL),
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

    # destaca a coluna Ação
    sug.conditional_formatting.add(
        f"G{cmp0}:G{r_cmp_tot-1}",
        FormulaRule(formula=[f'$G{cmp0}="Aumentar"'],
                    font=Font(color='1E7B34', bold=True),
                    fill=cor(VERDE)))
    sug.conditional_formatting.add(
        f"G{cmp0}:G{r_cmp_tot-1}",
        FormulaRule(formula=[f'$G{cmp0}="Reduzir"'],
                    font=vermelho, fill=fill_verm))

    # gráfico comparativo
    barra = BarChart()
    barra.type = 'col'
    barra.title = "Alocação atual x sugerida (%)"
    barra.height, barra.width = 9, 17
    barra.y_axis.numFmt = '0%'
    barra.x_axis.delete = False
    barra.y_axis.delete = False
    dref = Reference(sug, min_col=2, max_col=3,
                     min_row=R_CMP + 1, max_row=r_cmp_tot - 1)
    cref = Reference(sug, min_col=1, min_row=cmp0, max_row=r_cmp_tot - 1)
    barra.add_data(dref, titles_from_data=True)
    barra.set_categories(cref)
    sug.add_chart(barra, f"A{r_cmp_tot + 2}")

    # ─────────────────────────────────────────────────────────────────────
    # ABA 5 — INSTRUÇÕES
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
        "5. Aba 'Carteira Sugerida': à esquerda aparece a carteira atual por "
        "classe; à direita você monta manualmente a carteira recomendada "
        "(escolha a Classe, descreva o investimento e informe o valor).",
        "     O comparativo no fim mostra, classe a classe, o % atual x o % "
        "sugerido e quanto Aumentar ou Reduzir para chegar à carteira ideal.",
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
