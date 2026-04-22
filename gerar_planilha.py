#!/usr/bin/env python3
"""
Comparativo de Investimentos: Previdência vs CDB DI vs Fundo DI
Considera: Tabela regressiva IR, Come-cotas, Reaplicação de CDB
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────
# PARÂMETROS
# ─────────────────────────────────────────
APORTE_INICIAL   = 100_000.0   # R$
CDI_ANUAL        = 0.1075      # 10,75% a.a.
PRAZO_ANOS       = 10

CDB_PCT_CDI      = 1.00        # 100% do CDI
CDB_VENCIMENTO   = 2           # anos até vencimento (reaplicação)

FUNDO_TAXA_ADM   = 0.005       # 0,50% a.a.

PREV_TAXA_ADM    = 0.010       # 1,00% a.a.

# ─────────────────────────────────────────
# TABELAS DE IR
# ─────────────────────────────────────────

def ir_padrao(dias):
    if dias <= 180:  return 0.225
    elif dias <= 360: return 0.200
    elif dias <= 720: return 0.175
    else:            return 0.150

def ir_previdencia(anos):
    if anos <= 2:   return 0.35
    elif anos <= 4: return 0.30
    elif anos <= 6: return 0.25
    elif anos <= 8: return 0.20
    elif anos <= 10: return 0.15
    else:           return 0.10

def fator(taxa_anual, dias_uteis):
    return (1 + taxa_anual) ** (dias_uteis / 252)

# ─────────────────────────────────────────
# SIMULAÇÕES
# ─────────────────────────────────────────

def simular_cdb(aporte, cdi, pct_cdi, prazo_anos, venc_anos):
    """CDB com reaplicação no vencimento (reinicia alíquota)."""
    taxa = cdi * pct_cdi
    du_venc = int(venc_anos * 252)

    # lista de aplicações: [principal, dias_corridos]
    aplicacoes = [{'principal': aporte, 'dias': 0}]
    ir_pago_reaplicacao = 0.0
    resultados = []

    for ano in range(1, prazo_anos + 1):
        novas = []
        for app in aplicacoes:
            app['dias'] += 252
            app['valor'] = app['principal'] * fator(taxa, app['dias'])

        # verificar vencimentos
        for app in aplicacoes:
            if app['dias'] >= du_venc:
                ganho = app['valor'] - app['principal']
                ir = ganho * ir_padrao(app['dias'])
                ir_pago_reaplicacao += ir
                novas.append({'principal': app['valor'] - ir, 'dias': 0,
                               'valor': app['valor'] - ir})
            else:
                novas.append(app)
        aplicacoes = novas

        # resgate hipotético hoje
        liquido = 0.0
        for app in aplicacoes:
            ganho = app['valor'] - app['principal']
            ir = ganho * ir_padrao(app['dias'])
            liquido += app['valor'] - ir

        bruto_total = sum(a['valor'] for a in aplicacoes) + ir_pago_reaplicacao

        resultados.append({
            'ano': ano,
            'bruto': bruto_total,
            'liquido': liquido,
            'aliquota': ir_padrao(min(a['dias'] for a in aplicacoes)),
        })

    return resultados


def simular_fundo_di(aporte, cdi, taxa_adm, prazo_anos, longo_prazo=True):
    """Fundo DI com come-cotas semestral (maio e novembro)."""
    taxa_liq = cdi - taxa_adm
    aliq_comecotas = 0.15 if longo_prazo else 0.20

    valor = aporte
    base_comecotas = aporte     # referência do último come-cotas
    comecotas_pago = 0.0
    dias_desde_comecotas = 0
    DU_SEMESTRE = 126           # ~6 meses em dias úteis
    dias_total = 0
    resultados = []

    for ano in range(1, prazo_anos + 1):
        du_restantes = 252

        while du_restantes > 0:
            ate_prox = DU_SEMESTRE - dias_desde_comecotas
            avanco = min(ate_prox, du_restantes)

            valor *= fator(taxa_liq, avanco)
            dias_total += avanco
            dias_desde_comecotas += avanco
            du_restantes -= avanco

            if dias_desde_comecotas >= DU_SEMESTRE:
                ganho = valor - base_comecotas
                if ganho > 0:
                    ir = ganho * aliq_comecotas
                    comecotas_pago += ir
                    valor -= ir
                base_comecotas = valor
                dias_desde_comecotas = 0

        # resgate hipotético
        bruto_equiv = valor + comecotas_pago
        ganho_total = bruto_equiv - aporte
        ir_total_devido = ganho_total * ir_padrao(dias_total)
        ir_adicional = max(0.0, ir_total_devido - comecotas_pago)
        liquido = valor - ir_adicional

        resultados.append({
            'ano': ano,
            'bruto': bruto_equiv,
            'liquido': liquido,
            'comecotas': comecotas_pago,
            'aliquota': ir_padrao(dias_total),
        })

    return resultados


def simular_previdencia(aporte, cdi, taxa_adm, prazo_anos, tipo='VGBL'):
    """Previdência VGBL/PGBL: sem come-cotas, tabela regressiva própria."""
    taxa_liq = cdi - taxa_adm
    resultados = []

    for ano in range(1, prazo_anos + 1):
        bruto = aporte * fator(taxa_liq, ano * 252)
        aliq = ir_previdencia(ano)

        if tipo == 'PGBL':
            ir = bruto * aliq          # IR sobre o total
        else:
            ir = (bruto - aporte) * aliq  # IR só sobre rendimentos

        resultados.append({
            'ano': ano,
            'bruto': bruto,
            'liquido': bruto - ir,
            'aliquota': aliq,
        })

    return resultados


# ─────────────────────────────────────────
# GERAÇÃO DO EXCEL
# ─────────────────────────────────────────

def cor(hex_str):
    return PatternFill("solid", fgColor=hex_str)

def borda():
    lado = Side(style='thin', color='CCCCCC')
    return Border(left=lado, right=lado, top=lado, bottom=lado)

def fmt_brl(ws, cell):
    cell.number_format = 'R$ #,##0.00'

def estilo_header(cell, bg='1F4E79', fg='FFFFFF', negrito=True):
    cell.fill = cor(bg)
    cell.font = Font(bold=negrito, color=fg, size=11)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = borda()

def estilo_dado(cell, bg='FFFFFF'):
    cell.fill = cor(bg)
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = borda()
    cell.font = Font(size=10)


def criar_planilha():
    wb = openpyxl.Workbook()

    # ── Resultados das simulações ─────────────────────────────────────────────
    cdb_reap  = simular_cdb(APORTE_INICIAL, CDI_ANUAL, CDB_PCT_CDI,
                            PRAZO_ANOS, CDB_VENCIMENTO)
    cdb_livre = simular_cdb(APORTE_INICIAL, CDI_ANUAL, CDB_PCT_CDI,
                            PRAZO_ANOS, 999)  # sem vencimento forçado
    fundo_lp  = simular_fundo_di(APORTE_INICIAL, CDI_ANUAL, FUNDO_TAXA_ADM,
                                 PRAZO_ANOS, longo_prazo=True)
    fundo_cp  = simular_fundo_di(APORTE_INICIAL, CDI_ANUAL, FUNDO_TAXA_ADM,
                                 PRAZO_ANOS, longo_prazo=False)
    prev_vgbl = simular_previdencia(APORTE_INICIAL, CDI_ANUAL, PREV_TAXA_ADM,
                                    PRAZO_ANOS, 'VGBL')
    prev_pgbl = simular_previdencia(APORTE_INICIAL, CDI_ANUAL, PREV_TAXA_ADM,
                                    PRAZO_ANOS, 'PGBL')

    # ══════════════════════════════════════════════════════════════════════════
    # ABA 1 – PARÂMETROS
    # ══════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Parâmetros"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 18

    def add_titulo(row, texto, bg='1F4E79'):
        c = ws.cell(row=row, column=1, value=texto)
        c.fill = cor(bg)
        c.font = Font(bold=True, color='FFFFFF', size=12)
        c.alignment = Alignment(horizontal='left', vertical='center')
        c.border = borda()
        ws.merge_cells(f'A{row}:B{row}')
        ws.row_dimensions[row].height = 22

    def add_param(row, label, value, fmt=None, pct=False):
        lc = ws.cell(row=row, column=1, value=label)
        lc.fill = cor('EBF3FB')
        lc.font = Font(size=10)
        lc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        lc.border = borda()

        vc = ws.cell(row=row, column=2, value=value)
        vc.fill = cor('FFFFFF')
        vc.font = Font(bold=True, size=10, color='1F4E79')
        vc.alignment = Alignment(horizontal='right', vertical='center')
        vc.border = borda()
        if pct:
            vc.number_format = '0.00%'
        elif fmt:
            vc.number_format = fmt
        ws.row_dimensions[row].height = 20

    add_titulo(1, "⚙  PARÂMETROS DA SIMULAÇÃO")
    add_titulo(2, "Geral", bg='2E75B6')
    add_param(3,  "Aporte Inicial",             APORTE_INICIAL,    fmt='R$ #,##0.00')
    add_param(4,  "CDI Anual",                  CDI_ANUAL,         pct=True)
    add_param(5,  "Prazo da Simulação (anos)",  PRAZO_ANOS)

    add_titulo(7, "CDB DI", bg='2E75B6')
    add_param(8,  "% do CDI",                   CDB_PCT_CDI,       pct=True)
    add_param(9,  "Vencimento / Reaplicação (anos)", CDB_VENCIMENTO)

    add_titulo(11, "Fundo DI", bg='2E75B6')
    add_param(12, "Taxa de Administração",       FUNDO_TAXA_ADM,    pct=True)
    add_param(13, "Tipo (Longo Prazo)",          "Sim – come-cotas 15%")
    add_param(14, "Come-cotas (curto prazo)",    "20% – maio e novembro")

    add_titulo(16, "Previdência Privada (VGBL/PGBL)", bg='2E75B6')
    add_param(17, "Taxa de Administração",       PREV_TAXA_ADM,     pct=True)
    add_param(18, "VGBL – Base de IR",           "Somente rendimentos")
    add_param(19, "PGBL – Base de IR",           "Valor total resgatado")
    add_param(20, "Come-cotas",                  "NÃO")

    add_titulo(22, "TABELA REGRESSIVA – IR PADRÃO (CDB / Fundo)", bg='C55A11')
    for i, (faixa, aliq) in enumerate([
        ("Até 180 dias",       "22,5%"),
        ("181 a 360 dias",     "20,0%"),
        ("361 a 720 dias",     "17,5%"),
        ("Acima de 720 dias",  "15,0%"),
    ]):
        add_param(23 + i, faixa, aliq)

    add_titulo(28, "TABELA REGRESSIVA – PREVIDÊNCIA", bg='375623')
    for i, (faixa, aliq) in enumerate([
        ("Até 2 anos",         "35,0%"),
        ("2 a 4 anos",         "30,0%"),
        ("4 a 6 anos",         "25,0%"),
        ("6 a 8 anos",         "20,0%"),
        ("8 a 10 anos",        "15,0%"),
        ("Acima de 10 anos",   "10,0%"),
    ]):
        add_param(29 + i, faixa, aliq)

    # ══════════════════════════════════════════════════════════════════════════
    # ABA 2 – COMPARATIVO ANUAL
    # ══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Comparativo Anual")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = 'A3'

    colunas = [
        ("Ano", 7),
        ("CDB DI\nBruto", 16),
        ("CDB DI\nLíquido\n(c/ reaplicação)", 18),
        ("CDB DI\nLíquido\n(sem vencimento)", 18),
        ("Fundo DI LP\nBruto", 16),
        ("Fundo DI LP\nLíquido\n(após come-cotas)", 20),
        ("Fundo DI CP\nLíquido\n(come-cotas 20%)", 20),
        ("Prev. VGBL\nBruto", 16),
        ("Prev. VGBL\nLíquido", 16),
        ("Prev. PGBL\nLíquido", 16),
    ]

    CORES_HEADER = ['1F4E79','2E75B6','2E75B6','4472C4',
                    'C55A11','E2750C','E2750C',
                    '375623','548235','548235']

    for col, (titulo, largura) in enumerate(colunas, 1):
        ws2.column_dimensions[get_column_letter(col)].width = largura
        c = ws2.cell(row=1, column=col, value=titulo)
        estilo_header(c, bg=CORES_HEADER[col - 1])

    # linha de referência: aporte
    ws2.row_dimensions[1].height = 52
    ws2.row_dimensions[2].height = 18
    ref = ws2.cell(row=2, column=1, value="Aporte")
    ref.font = Font(bold=True, size=10)
    ref.alignment = Alignment(horizontal='center')
    for col in range(2, 11):
        c = ws2.cell(row=2, column=col, value=APORTE_INICIAL)
        c.number_format = 'R$ #,##0.00'
        c.font = Font(bold=True, italic=True, color='555555', size=10)
        c.alignment = Alignment(horizontal='right')

    ZEBRA_A = 'EBF3FB'
    ZEBRA_B = 'FFFFFF'

    for i in range(PRAZO_ANOS):
        row = i + 3
        bg = ZEBRA_A if i % 2 == 0 else ZEBRA_B
        ws2.row_dimensions[row].height = 18

        dados = [
            i + 1,
            cdb_reap[i]['bruto'],
            cdb_reap[i]['liquido'],
            cdb_livre[i]['liquido'],
            fundo_lp[i]['bruto'],
            fundo_lp[i]['liquido'],
            fundo_cp[i]['liquido'],
            prev_vgbl[i]['bruto'],
            prev_vgbl[i]['liquido'],
            prev_pgbl[i]['liquido'],
        ]

        for col, val in enumerate(dados, 1):
            c = ws2.cell(row=row, column=col, value=val)
            estilo_dado(c, bg=bg)
            if col == 1:
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.font = Font(bold=True, size=10)
            else:
                c.number_format = 'R$ #,##0.00'

    # ══════════════════════════════════════════════════════════════════════════
    # ABA 3 – ALÍQUOTAS E IR
    # ══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Alíquotas e IR")
    ws3.sheet_view.showGridLines = False

    colunas3 = [
        ("Ano", 7),
        ("CDB\nAlíquota", 14),
        ("CDB c/ Reapl.\nIR pago total", 16),
        ("Fundo LP\nAlíquota", 14),
        ("Fundo LP\nCome-cotas\nacumulado", 18),
        ("Prev. VGBL\nAlíquota", 14),
        ("Prev. VGBL\nIR no resgate", 16),
        ("Prev. PGBL\nAlíquota", 14),
        ("Prev. PGBL\nIR no resgate", 16),
    ]

    CORES_H3 = ['1F4E79','2E75B6','2E75B6','C55A11','C55A11',
                '375623','375623','375623','375623']

    for col, (titulo, largura) in enumerate(colunas3, 1):
        ws3.column_dimensions[get_column_letter(col)].width = largura
        c = ws3.cell(row=1, column=col, value=titulo)
        estilo_header(c, bg=CORES_H3[col - 1])
    ws3.row_dimensions[1].height = 52

    for i in range(PRAZO_ANOS):
        row = i + 2
        bg = ZEBRA_A if i % 2 == 0 else ZEBRA_B
        ws3.row_dimensions[row].height = 18

        ano = i + 1
        aliq_cdb = ir_padrao(min(ano * 252, CDB_VENCIMENTO * 252))
        ir_cdb_pago = cdb_reap[i]['bruto'] - cdb_reap[i]['liquido']  # approx
        aliq_fundo = fundo_lp[i]['aliquota']
        cc_fundo = fundo_lp[i]['comecotas']
        aliq_vgbl = prev_vgbl[i]['aliquota']
        ir_vgbl = prev_vgbl[i]['bruto'] - prev_vgbl[i]['liquido']
        aliq_pgbl = prev_pgbl[i]['aliquota']
        ir_pgbl = prev_pgbl[i]['bruto'] - prev_pgbl[i]['liquido']

        linha = [ano, aliq_cdb, ir_cdb_pago, aliq_fundo, cc_fundo,
                 aliq_vgbl, ir_vgbl, aliq_pgbl, ir_pgbl]

        for col, val in enumerate(linha, 1):
            c = ws3.cell(row=row, column=col, value=val)
            estilo_dado(c, bg=bg)
            if col == 1:
                c.alignment = Alignment(horizontal='center')
                c.font = Font(bold=True, size=10)
            elif col in (2, 4, 6, 8):
                c.number_format = '0.0%'
            else:
                c.number_format = 'R$ #,##0.00'

    # ══════════════════════════════════════════════════════════════════════════
    # ABA 4 – GRÁFICO VALOR LÍQUIDO
    # ══════════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("Gráfico – Valor Líquido")

    chart = LineChart()
    chart.title = "Evolução do Valor Líquido por Produto"
    chart.style = 10
    chart.y_axis.title = "R$"
    chart.x_axis.title = "Ano"
    chart.width  = 28
    chart.height = 18
    chart.y_axis.numFmt = 'R$ #,##0'

    # Referência de dados do Comparativo Anual (colunas 3, 4, 6, 9, 10)
    for col_idx, nome in [
        (3, "CDB DI (c/ reaplicação)"),
        (4, "CDB DI (sem vencimento)"),
        (6, "Fundo DI LP"),
        (7, "Fundo DI CP"),
        (9, "Prev. VGBL"),
        (10, "Prev. PGBL"),
    ]:
        dados_ref = Reference(ws2, min_col=col_idx, min_row=3,
                              max_row=2 + PRAZO_ANOS)
        chart.add_data(dados_ref, titles_from_data=False)
        from openpyxl.chart.series import SeriesLabel
        chart.series[-1].title = SeriesLabel(v=nome)

    cats = Reference(ws2, min_col=1, min_row=3, max_row=2 + PRAZO_ANOS)
    chart.set_categories(cats)
    chart.shape = 4
    ws4.add_chart(chart, "A2")

    # ══════════════════════════════════════════════════════════════════════════
    # ABA 5 – RESUMO FINAL
    # ══════════════════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet("Resumo Final")
    ws5.sheet_view.showGridLines = False

    titulo_resumo = ws5.cell(row=1, column=1,
        value=f"RESUMO FINAL – {PRAZO_ANOS} ANOS  |  Aporte: R$ {APORTE_INICIAL:,.2f}  |  CDI: {CDI_ANUAL*100:.2f}% a.a.")
    titulo_resumo.fill = cor('1F4E79')
    titulo_resumo.font = Font(bold=True, color='FFFFFF', size=13)
    titulo_resumo.alignment = Alignment(horizontal='center', vertical='center')
    ws5.merge_cells('A1:F1')
    ws5.row_dimensions[1].height = 28

    headers5 = ["Produto", "Valor Bruto Final", "Valor Líquido Final",
                "Rendimento Líquido", "Rentab. Líq. (%)", "Observação"]
    widths5   = [30, 20, 20, 20, 18, 40]

    for col, (h, w) in enumerate(zip(headers5, widths5), 1):
        ws5.column_dimensions[get_column_letter(col)].width = w
        c = ws5.cell(row=2, column=col, value=h)
        estilo_header(c, bg='2E75B6')
    ws5.row_dimensions[2].height = 22

    produtos = [
        ("CDB DI (c/ reaplicação a cada 2 anos)",
         cdb_reap[-1]['bruto'], cdb_reap[-1]['liquido'],
         f"Reaplicação a cada {CDB_VENCIMENTO} anos – reinicia alíquota"),
        ("CDB DI (sem vencimento forçado)",
         cdb_livre[-1]['bruto'], cdb_livre[-1]['liquido'],
         "Mantido até o final – alíquota 15% (>720 dias)"),
        ("Fundo DI – Longo Prazo",
         fundo_lp[-1]['bruto'], fundo_lp[-1]['liquido'],
         f"Come-cotas 15% (mai/nov) | Taxa adm: {FUNDO_TAXA_ADM*100:.2f}% a.a."),
        ("Fundo DI – Curto Prazo",
         fundo_cp[-1]['bruto'], fundo_cp[-1]['liquido'],
         f"Come-cotas 20% (mai/nov) | Taxa adm: {FUNDO_TAXA_ADM*100:.2f}% a.a."),
        ("Previdência VGBL",
         prev_vgbl[-1]['bruto'], prev_vgbl[-1]['liquido'],
         f"IR só rendimentos | Taxa adm: {PREV_TAXA_ADM*100:.2f}% | Alíq. 10%+ de 10 anos"),
        ("Previdência PGBL",
         prev_pgbl[-1]['bruto'], prev_pgbl[-1]['liquido'],
         f"IR sobre total | Taxa adm: {PREV_TAXA_ADM*100:.2f}% | Indicado p/ quem deduz IR"),
    ]

    CORES5 = ['D6E4F0','D6E4F0','FDEBD0','FDEBD0','D5F5E3','D5F5E3']

    for i, (nome, bruto, liq, obs) in enumerate(produtos):
        row = i + 3
        bg = CORES5[i]
        ws5.row_dimensions[row].height = 22
        rend = liq - APORTE_INICIAL
        pct  = rend / APORTE_INICIAL

        linha = [nome, bruto, liq, rend, pct, obs]
        for col, val in enumerate(linha, 1):
            c = ws5.cell(row=row, column=col, value=val)
            c.fill = cor(bg)
            c.border = borda()
            c.alignment = Alignment(horizontal='right' if col > 1 else 'left',
                                    vertical='center', indent=1)
            c.font = Font(size=10)
            if col == 2: c.number_format = 'R$ #,##0.00'
            elif col == 3:
                c.number_format = 'R$ #,##0.00'
                c.font = Font(bold=True, size=10)
            elif col == 4: c.number_format = 'R$ #,##0.00'
            elif col == 5: c.number_format = '0.00%'

    # destaque do melhor líquido
    melhor_row = 3 + max(range(len(produtos)), key=lambda i: produtos[i][2])
    for col in range(1, 7):
        c = ws5.cell(row=melhor_row, column=col)
        c.fill = cor('FFD700')
        c.font = Font(bold=True, size=10)

    # nota
    nota_row = len(produtos) + 5
    nota = ws5.cell(row=nota_row, column=1,
        value="* Linha destacada em dourado = melhor rentabilidade líquida no período simulado.")
    nota.font = Font(italic=True, size=9, color='555555')
    ws5.merge_cells(f'A{nota_row}:F{nota_row}')

    nota2 = ws5.cell(row=nota_row+1, column=1,
        value="* PGBL vantajoso somente para quem declara IR completo e tem dedução na fase de acumulação (economia de até 27,5% no aporte).")
    nota2.font = Font(italic=True, size=9, color='555555')
    ws5.merge_cells(f'A{nota_row+1}:F{nota_row+1}')

    # ── Salvar ────────────────────────────────────────────────────────────────
    wb.active = ws5   # abre no resumo
    path = "/home/user/Planilha/comparativo_investimentos.xlsx"
    wb.save(path)
    print(f"Planilha gerada: {path}")


if __name__ == "__main__":
    criar_planilha()
