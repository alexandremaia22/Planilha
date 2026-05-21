#!/usr/bin/env python3
"""
Converte um .xlsx (gerado pelo openpyxl, com gráficos) em um .xlsm com:
  - o projeto VBA (vbaProject.bin) contendo a macro AdicionarLinha;
  - um botão "+ Adicionar linha de ativo" na aba, ligado à macro.
"""

import zipfile
from vba_builder import build as build_vba

# ── código da macro ──────────────────────────────────────────────────────
MACRO = (
    'Sub AdicionarLinha()\n'
    "    ' Insere uma nova linha de ativo logo abaixo da celula selecionada,\n"
    "    ' copiando as formulas de rentabilidade e limpando os campos de "
    "entrada.\n"
    '    Dim r As Long\n'
    '    r = ActiveCell.Row\n'
    '    If r < 10 Then\n'
    '        MsgBox "Selecione uma celula em uma linha de ativo (abaixo do "'
    ' & "cabecalho de uma classe) antes de adicionar a linha.", _\n'
    '               vbExclamation, "Adicionar linha"\n'
    '        Exit Sub\n'
    '    End If\n'
    '    Application.ScreenUpdating = False\n'
    '    Rows(r).Copy\n'
    '    Rows(r + 1).Insert Shift:=xlDown, '
    'CopyOrigin:=xlFormatFromLeftOrAbove\n'
    '    Application.CutCopyMode = False\n'
    '    Dim nova As Long, col As Long\n'
    '    nova = r + 1\n'
    '    Rows(nova).EntireRow.Hidden = False\n'
    '    Cells(nova, 1).Value = ""\n'
    '    Cells(nova, 2).ClearContents\n'
    '    For col = 3 To 58 Step 5\n'
    '        Cells(nova, col).ClearContents\n'
    '        Cells(nova, col + 1).ClearContents\n'
    '        Cells(nova, col + 2).ClearContents\n'
    '    Next col\n'
    '    Cells(nova, 1).Select\n'
    '    Application.ScreenUpdating = True\n'
    'End Sub\n'
)

# ── XML do botão (shape com macro associada) ──────────────────────────────
# Inserido no drawing da aba; usa o namespace padrão (spreadsheetDrawing) e
# o prefixo a: (drawingml), ambos já declarados pelo openpyxl no <wsDr>.
BOTAO_XML = (
    '<twoCellAnchor editAs="oneCell">'
    '<from><col>7</col><colOff>19050</colOff>'
    '<row>1</row><rowOff>9525</rowOff></from>'
    '<to><col>10</col><colOff>0</colOff><row>3</row><rowOff>0</rowOff></to>'
    '<sp macro="[0]!AdicionarLinha" textlink="">'
    '<nvSpPr>'
    '<cNvPr id="500" name="BotaoAdicionarLinha"/>'
    '<cNvSpPr/>'
    '</nvSpPr>'
    '<spPr>'
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
    '<a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill>'
    '<a:ln w="9525"><a:solidFill><a:srgbClr val="14365C"/></a:solidFill>'
    '</a:ln>'
    '</spPr>'
    '<txBody>'
    '<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" '
    'lIns="36000" tIns="18000" rIns="36000" bIns="18000" anchor="ctr"/>'
    '<a:lstStyle/>'
    '<a:p><a:pPr algn="ctr"/>'
    '<a:r><a:rPr lang="pt-BR" sz="1000" b="1">'
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
    '<a:t>+ Adicionar linha de ativo</a:t></a:r>'
    '</a:p>'
    '</txBody>'
    '</sp>'
    '<clientData/>'
    '</twoCellAnchor>'
)

WB_CT_XLSX = ('application/vnd.openxmlformats-officedocument.'
              'spreadsheetml.sheet.main+xml')
WB_CT_XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'


def converter(xlsx_path: str, xlsm_path: str) -> None:
    with zipfile.ZipFile(xlsx_path) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    if 'xl/drawings/drawing1.xml' not in items:
        raise RuntimeError("o .xlsx precisa conter o drawing dos gráficos")

    # 1. content types: workbook macro-enabled + parte vbaProject
    ct = items['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace(WB_CT_XLSX, WB_CT_XLSM)
    ct = ct.replace(
        '</Types>',
        '<Override PartName="/xl/vbaProject.bin" '
        'ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    items['[Content_Types].xml'] = ct.encode('utf-8')

    # 2. relacionamento do workbook -> vbaProject.bin
    rels = items['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rels = rels.replace(
        '</Relationships>',
        f'<Relationship Id="rIdVBA" Type="{VBA_REL}" '
        'Target="vbaProject.bin"/></Relationships>')
    items['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')

    # 3. botão no drawing da aba
    dr = items['xl/drawings/drawing1.xml'].decode('utf-8')
    dr = dr.replace('</wsDr>', BOTAO_XML + '</wsDr>')
    items['xl/drawings/drawing1.xml'] = dr.encode('utf-8')

    # 4. o próprio projeto VBA
    items['xl/vbaProject.bin'] = build_vba(MACRO)

    with zipfile.ZipFile(xlsm_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
