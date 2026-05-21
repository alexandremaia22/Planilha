#!/usr/bin/env python3
"""
Gera um vbaProject.bin minimo e valido (formato OLE/CFB + MS-OVBA) contendo
um unico modulo de codigo VBA. Usado para embutir a macro de "Adicionar linha"
na planilha de acompanhamento.
"""

import struct

ENDOFCHAIN = 0xFFFFFFFE
FREESECT   = 0xFFFFFFFF
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF

SECTOR      = 512
MINISECTOR  = 64
MINICUTOFF  = 4096

PROJECT_GUID = "{5DD90D76-4904-47A2-AF0D-D69B4673F103}"


# ─────────────────────────────────────────────────────────────────────────
# Compressao MS-OVBA (apenas tokens literais — sempre valida e exata)
# ─────────────────────────────────────────────────────────────────────────
def _compress(data: bytes) -> bytes:
    """CompressedContainer com um unico CompressedChunk so de literais."""
    if len(data) > 3640:
        raise ValueError("conteudo grande demais para um unico chunk literal")
    tokens = bytearray()
    for i in range(0, len(data), 8):
        tokens.append(0x00)              # FlagByte: 8 literais
        tokens += data[i:i + 8]
    size_field = (len(tokens) + 2) - 3   # CompressedChunkSize - 3
    header = 0x8000 | 0x3000 | (size_field & 0x0FFF)
    return b'\x01' + struct.pack('<H', header) + bytes(tokens)


def _decompress(container: bytes) -> bytes:
    """Descompressor de verificacao (suporta chunks comprimidos e nao)."""
    assert container[0] == 0x01
    out = bytearray()
    pos = 1
    while pos < len(container):
        header = struct.unpack('<H', container[pos:pos + 2])[0]
        pos += 2
        size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        end = pos + size - 2
        if not compressed:
            out += container[pos:pos + 4096]
            pos += 4096
            continue
        while pos < end:
            flags = container[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(container[pos])
                    pos += 1
                else:
                    token = struct.unpack('<H', container[pos:pos + 2])[0]
                    pos += 2
                    diff = len(out) - 0
                    bits = max(4, (diff - 1).bit_length())
                    length = (token & ((1 << (16 - bits)) - 1)) + 3
                    offset = (token >> (16 - bits)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
    return bytes(out)


# ─────────────────────────────────────────────────────────────────────────
# Stream "dir" (descomprimido)
# ─────────────────────────────────────────────────────────────────────────
def _rec(rid: int, data: bytes) -> bytes:
    return struct.pack('<HI', rid, len(data)) + data


def _build_dir(module_name: str, project_name: str) -> bytes:
    mn = module_name.encode('cp1252')
    mnu = module_name.encode('utf-16-le')
    pn = project_name.encode('cp1252')

    d = bytearray()
    # PROJECTINFORMATION
    d += _rec(0x0001, struct.pack('<I', 0x00000001))      # SYSKIND  (Win32)
    d += _rec(0x0002, struct.pack('<I', 0x00000409))      # LCID
    d += _rec(0x0014, struct.pack('<I', 0x00000409))      # LCIDINVOKE
    d += _rec(0x0003, struct.pack('<H', 0x04E4))          # CODEPAGE (1252)
    d += _rec(0x0004, pn)                                 # NAME
    d += _rec(0x0005, b'')                                # DOCSTRING
    d += _rec(0x0040, b'')                                #   unicode
    d += _rec(0x0006, b'')                                # HELPFILEPATH 1
    d += _rec(0x003D, b'')                                #   helpfile 2
    d += _rec(0x0007, struct.pack('<I', 0))               # HELPCONTEXT
    d += _rec(0x0008, struct.pack('<I', 0))               # LIBFLAGS
    d += struct.pack('<HIIH', 0x0009, 0x00000004, 1, 0)   # VERSION
    d += _rec(0x000C, b'')                                # CONSTANTS
    d += _rec(0x003C, b'')                                #   unicode
    # PROJECTREFERENCES: nenhuma
    # PROJECTMODULES
    d += _rec(0x000F, struct.pack('<H', 1))               # MODULES count
    d += _rec(0x0013, struct.pack('<H', 0xFFFF))          # PROJECTCOOKIE
    d += _rec(0x0019, mn)                                 # MODULENAME
    d += _rec(0x0047, mnu)                                # MODULENAMEUNICODE
    d += _rec(0x001A, mn)                                 # MODULESTREAMNAME
    d += _rec(0x0032, mnu)                                #   unicode
    d += _rec(0x001C, b'')                                # MODULEDOCSTRING
    d += _rec(0x0048, b'')                                #   unicode
    d += _rec(0x0031, struct.pack('<I', 0))               # MODULEOFFSET
    d += _rec(0x001E, struct.pack('<I', 0))               # MODULEHELPCONTEXT
    d += _rec(0x002C, struct.pack('<H', 0xFFFF))          # MODULECOOKIE
    d += _rec(0x0021, b'')                                # MODULETYPE (proc)
    d += _rec(0x002B, b'')                                # module terminator
    # terminador do dir
    d += struct.pack('<HI', 0x0010, 0x00000000)
    return bytes(d)


# ─────────────────────────────────────────────────────────────────────────
# Streams de texto
# ─────────────────────────────────────────────────────────────────────────
def _build_project(module_name: str, project_name: str) -> bytes:
    # Projeto sem protecao: CMG/DPB/GC omitidos (o Excel os recria ao salvar).
    txt = (
        f'ID="{PROJECT_GUID}"\r\n'
        f'Module={module_name}\r\n'
        f'Name="{project_name}"\r\n'
        f'HelpContextID="0"\r\n'
        f'VersionCompatible32="393222000"\r\n'
        '\r\n[Host Extender Info]\r\n'
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n'
        '\r\n[Workspace]\r\n'
        f'{module_name}=0, 0, 0, 0, C \r\n'
    )
    return txt.encode('cp1252')


def _build_projectwm(module_name: str) -> bytes:
    return (module_name.encode('cp1252') + b'\x00'
            + module_name.encode('utf-16-le') + b'\x00\x00'
            + b'\x00\x00')


def _build_vba_project() -> bytes:
    # Reserved1=0x61CC, Version, Reserved2=0x00, Reserved3, PerformanceCache
    return b'\xCC\x61\xFF\xFF\x00\x00\x00'


# ─────────────────────────────────────────────────────────────────────────
# Escrita do contêiner OLE/CFB
# ─────────────────────────────────────────────────────────────────────────
def _dir_entry(name, objtype, color, left, right, child,
               start, size):
    raw = name.encode('utf-16-le') + b'\x00\x00'
    raw = raw[:64].ljust(64, b'\x00')
    namelen = (len(name) + 1) * 2
    return (raw
            + struct.pack('<H', namelen)
            + struct.pack('<B', objtype)
            + struct.pack('<B', color)
            + struct.pack('<I', left)
            + struct.pack('<I', right)
            + struct.pack('<I', child)
            + b'\x00' * 16                     # CLSID
            + struct.pack('<I', 0)             # state
            + b'\x00' * 8                      # ctime
            + b'\x00' * 8                      # mtime
            + struct.pack('<I', start)
            + struct.pack('<Q', size))


def build(module_source: str,
          module_name: str = "Macros",
          project_name: str = "VBAProject") -> bytes:
    src = ('Attribute VB_Name = "%s"\r\n' % module_name
           + module_source.replace('\n', '\r\n'))
    src_bytes = src.encode('cp1252')

    module_stream = _compress(src_bytes)
    assert _decompress(module_stream) == src_bytes, "round-trip do modulo"

    dir_raw = _build_dir(module_name, project_name)
    dir_stream = _compress(dir_raw)
    assert _decompress(dir_stream) == dir_raw, "round-trip do dir"

    project_stream   = _build_project(module_name, project_name)
    projectwm_stream = _build_projectwm(module_name)
    vbaproj_stream   = _build_vba_project()

    # ── streams (todos pequenos -> mini stream) ──────────────────────────
    # ordem no mini stream
    streams = [
        ("PROJECT",      project_stream),
        ("PROJECTwm",    projectwm_stream),
        ("_VBA_PROJECT", vbaproj_stream),
        ("dir",          dir_stream),
        ("Macros",       module_stream),
    ]

    mini = bytearray()
    mini_info = {}          # nome -> (start_minisector, size)
    for name, data in streams:
        start = len(mini) // MINISECTOR
        mini_info[name] = (start, len(data))
        mini += data
        while len(mini) % MINISECTOR:
            mini.append(0)
    n_minisectors = len(mini) // MINISECTOR

    # ── miniFAT ──────────────────────────────────────────────────────────
    minifat = [FREESECT] * (SECTOR // 4)
    for name, data in streams:
        start, size = mini_info[name]
        count = (size + MINISECTOR - 1) // MINISECTOR
        for k in range(count):
            ms = start + k
            minifat[ms] = ENDOFCHAIN if k == count - 1 else ms + 1

    # ── layout dos setores regulares ─────────────────────────────────────
    # 0: FAT | 1,2: diretorio | 3: miniFAT | 4..: mini stream
    mini_padded = bytes(mini).ljust(
        ((len(mini) + SECTOR - 1) // SECTOR) * SECTOR, b'\x00')
    n_mini_sec = len(mini_padded) // SECTOR
    first_mini_sec = 4

    # ── FAT ──────────────────────────────────────────────────────────────
    fat = [FREESECT] * (SECTOR // 4)
    fat[0] = FATSECT
    fat[1] = 2
    fat[2] = ENDOFCHAIN                       # diretorio
    fat[3] = ENDOFCHAIN                       # miniFAT
    for k in range(n_mini_sec):
        s = first_mini_sec + k
        fat[s] = ENDOFCHAIN if k == n_mini_sec - 1 else s + 1

    # ── diretorio ────────────────────────────────────────────────────────
    # DIDs: 0 Root,1 VBA,2 PROJECT,3 PROJECTwm,4 dir,5 Macros,6 _VBA_PROJECT
    entries = []
    entries.append(_dir_entry("Root Entry", 5, 1,
                              NOSTREAM, NOSTREAM, 2,
                              first_mini_sec, len(mini)))
    entries.append(_dir_entry("VBA", 1, 0,
                              NOSTREAM, NOSTREAM, 5, ENDOFCHAIN, 0))
    pj = mini_info["PROJECT"]
    entries.append(_dir_entry("PROJECT", 2, 1, 1, 3, NOSTREAM, pj[0], pj[1]))
    pw = mini_info["PROJECTwm"]
    entries.append(_dir_entry("PROJECTwm", 2, 0,
                              NOSTREAM, NOSTREAM, NOSTREAM, pw[0], pw[1]))
    dr = mini_info["dir"]
    entries.append(_dir_entry("dir", 2, 0,
                              NOSTREAM, NOSTREAM, NOSTREAM, dr[0], dr[1]))
    mc = mini_info["Macros"]
    entries.append(_dir_entry("Macros", 2, 1, 4, 6, NOSTREAM, mc[0], mc[1]))
    vp = mini_info["_VBA_PROJECT"]
    entries.append(_dir_entry("_VBA_PROJECT", 2, 0,
                              NOSTREAM, NOSTREAM, NOSTREAM, vp[0], vp[1]))
    while len(entries) % 4:                   # completa o setor
        entries.append(b'\x00' * 128)
    directory = b''.join(entries)
    n_dir_sec = len(directory) // SECTOR      # = 2

    # ── cabecalho ────────────────────────────────────────────────────────
    header = bytearray(512)
    header[0:8]   = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    header[8:24]  = b'\x00' * 16
    struct.pack_into('<H', header, 24, 0x003E)        # minor
    struct.pack_into('<H', header, 26, 0x0003)        # major (v3)
    struct.pack_into('<H', header, 28, 0xFFFE)        # byte order
    struct.pack_into('<H', header, 30, 0x0009)        # sector shift (512)
    struct.pack_into('<H', header, 32, 0x0006)        # mini sector shift
    struct.pack_into('<I', header, 40, 0)             # num dir sectors (v3=0)
    struct.pack_into('<I', header, 44, 1)             # num FAT sectors
    struct.pack_into('<I', header, 48, 1)             # first dir sector
    struct.pack_into('<I', header, 52, 0)             # transaction sig
    struct.pack_into('<I', header, 56, MINICUTOFF)    # mini stream cutoff
    struct.pack_into('<I', header, 60, 3)             # first miniFAT sector
    struct.pack_into('<I', header, 64, 1)             # num miniFAT sectors
    struct.pack_into('<I', header, 68, ENDOFCHAIN)    # first DIFAT sector
    struct.pack_into('<I', header, 72, 0)             # num DIFAT sectors
    difat = [0] + [FREESECT] * 108
    for i, v in enumerate(difat):
        struct.pack_into('<I', header, 76 + i * 4, v)

    # ── montagem ─────────────────────────────────────────────────────────
    out = bytearray()
    out += header
    out += struct.pack('<%dI' % len(fat), *fat)       # setor 0: FAT
    out += directory                                  # setores 1-2
    out += struct.pack('<%dI' % len(minifat), *minifat)  # setor 3: miniFAT
    out += mini_padded                                # setores 4..
    return bytes(out)


if __name__ == '__main__':
    macro = (
        'Sub AdicionarLinha()\n'
        '    Dim r As Long\n'
        '    r = ActiveCell.Row\n'
        '    If r < 10 Then\n'
        '        MsgBox "Selecione uma celula dentro de uma classe (linha de '
        'ativo) antes de adicionar."\n'
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
    data = build(macro)
    with open('vbaProject.bin', 'wb') as fh:
        fh.write(data)
    print("vbaProject.bin gerado:", len(data), "bytes")
