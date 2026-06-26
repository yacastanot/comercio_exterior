"""
Migración a Python: Importaciones NIT 2026 – Combustibles (CUCIsubg 3341/3343)
Equivalente al programa SAS 'Impo NIT 2026.sas'.

Procesa tres años (2024-2025-2026), agrupa por (posara, mes, paor, regimen, NIT, RAZON)
y conserva solo productos con CUCIsubg 3341 ó 3343 (combustibles) tras cruzar con corre62.
Dependencias: pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

RUTA_ASU       = Path(r'\\systema44\Migracion\e100')
RUTA_COMEX     = Path(r'D:\COMEX')
RUTA_RESULTADO = RUTA_COMEX / 'IMPO' / 'Resultados'
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Importaciones 2026_NIT.xlsx'
CORRE62_XLSX   = RUTA_COMEX / 'corre62.xlsx'

# 2024: año completo (12 meses)
ARCHIVOS_2024 = [
    RUTA_ASU / 'M10124.asu', RUTA_ASU / 'M10224.asu',
    RUTA_ASU / 'M10324.asu', RUTA_ASU / 'M10424.asu',
    RUTA_ASU / 'M10524.asu', RUTA_ASU / 'M10624.asu',
    RUTA_ASU / 'M10724.asu', RUTA_ASU / 'M10824.asu',
    RUTA_ASU / 'M10924.asu', RUTA_ASU / 'M11024.asu',
    RUTA_ASU / 'M11124.asu', RUTA_ASU / 'M11224.asu',
]

ARCHIVOS_2025 = [
    RUTA_ASU / 'M10125.asu', RUTA_ASU / 'M10225.asu',
    RUTA_ASU / 'M10325.asu', RUTA_ASU / 'M10425.asu',
    RUTA_ASU / 'M10525.asu', RUTA_ASU / 'M10625.asu',
    RUTA_ASU / 'M10725.asu', RUTA_ASU / 'M10825.asu',
    RUTA_ASU / 'M10925.asu', RUTA_ASU / 'M11025.asu',
    RUTA_ASU / 'M11125.asu', RUTA_ASU / 'M11225.asu',
]

ARCHIVOS_2026 = [
    RUTA_ASU / 'M10126.asu',
    RUTA_ASU / 'M10226.asu',
    # RUTA_ASU / 'M10326.asu',
    # RUTA_ASU / 'M10426.asu',
    # RUTA_ASU / 'M10526.asu',
    # RUTA_ASU / 'M10626.asu',
    # RUTA_ASU / 'M10726.asu',
    # RUTA_ASU / 'M10826.asu',
    # RUTA_ASU / 'M10926.asu',
    # RUTA_ASU / 'M11026.asu',
    # RUTA_ASU / 'M11126.asu',
    # RUTA_ASU / 'M11226.asu',
]

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO (LRECL=450)
# =============================================================================

COLSPECS = [
    (0,  4),    # FECH     @1  $4.
    (4,  6),    # ADUA     @5  2.
    (6,  9),    # PAOR     @7  3.
    (9,  12),   # PAPR     @10 3.
    (12, 15),   # PACO     @13 3.
    (15, 17),   # DEPTO    @16 2.
    (17, 19),   # VIA      @18 2.
    (19, 22),   # BAND     @20 3.
    (22, 26),   # REGIMEN  $23-26
    (26, 29),   # ACUERDO  @27 $3.
    (29, 42),   # PBK      @30 13.2
    (42, 55),   # KNETO    @43 13.2
    (55, 68),   # CANU     @56 13.2
    (68, 71),   # CODUN    @69 $3.
    (71, 89),   # POSARA   @72 18.
    (89, 100),  # FOBDO    @90 11.2
    (100, 111), # FLET     @101 11.2
    (111, 124), # CIFDO    @112 13.2
    (124, 139), # CIFPE    @125 15.
    (160, 189), # CIUDAD   @161 29.
    (226, 230), # ACTIVID  @227 $4.
    (304, 317), # SEGUROS  @305 13.2
    (317, 330), # OTROSG   @318 13.2
    (363, 379), # NIT      @364 $16.
    (379, 380), # DIGV     @380 $1.
    (380, 440), # RAZON    @381 $60.
]

NOMBRES = [
    'FECH', 'ADUA', 'PAOR', 'PAPR', 'PACO', 'DEPTO', 'VIA', 'BAND',
    'REGIMEN', 'ACUERDO', 'PBK', 'KNETO', 'CANU', 'CODUN', 'POSARA',
    'FOBDO', 'FLET', 'CIFDO', 'CIFPE', 'CIUDAD', 'ACTIVID',
    'SEGUROS', 'OTROSG', 'NIT', 'DIGV', 'RAZON',
]

COLS_STR = {
    'FECH', 'PAPR', 'PACO', 'REGIMEN', 'ACUERDO', 'CODUN',
    'POSARA', 'CIUDAD', 'ACTIVID', 'NIT', 'DIGV', 'RAZON',
}

COLS_DEC2 = ['PBK', 'KNETO', 'CANU', 'FOBDO', 'FLET', 'CIFDO', 'SEGUROS', 'OTROSG']

REGIMEN_EXCLUIR = {
    'C390', 'C392', 'C393', 'C394', 'C395', 'C396', 'C397',
    'C540', 'C541', 'C545', 'C546', 'C660', 'C665',
    'S100', 'S105', 'S106', 'S200', 'S140', 'S240', 'C398', 'C547',
}

NITS_AEREAS = {
    '0000000890912462', '000000890912462', '00000890912462', '0000890912462',
    '000890912462', '00890912462', '0890912462', '890912462',
    '000000000890100577', '00000000890100577', '0000000890100577',
    '000000890100577', '00000890100577', '0000890100577',
    '000890100577', '00890100577', '0890100577', '890100577',
}
POSARA_AEREAS  = {'8802400000', '8802309000'}
REGIMEN_AEREAS = {'C190', 'C196'}

# Agrupación: equivale a %CIF(posara mes paor regimen NIT RAZON)
GRUPO     = ['POSARA', 'MES', 'PAOR', 'REGIMEN', 'NIT', 'RAZON']
VARS_SUMA = ['CIFDO', 'CIFPE', 'FOBDO', 'KNETO', 'PBK', 'CANU']

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    def _conv(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        return float(s) if '.' in s else (float(s) / 100 if s else None)
    return serie.apply(_conv)


def leer_asu(archivos: list) -> pd.DataFrame:
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            ruta,
            colspecs=COLSPECS,
            names=NOMBRES,
            dtype={c: str for c in COLS_STR},
            header=None,
            encoding='latin-1',
        )
        partes.append(df)
        print(f"  Leído: {Path(ruta).name}  ({len(df):,} registros)")

    df = pd.concat(partes, ignore_index=True)

    # Strip whitespace antes de la conversión numérica (PAOR puede ser aún object)
    for col in ('REGIMEN', 'PAOR', 'NIT', 'RAZON', 'ACUERDO'):
        if df[col].dtype == object:
            df[col] = df[col].str.strip()

    for col in NOMBRES:
        if col not in COLS_STR:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for col in COLS_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    df['POSARA']   = df['POSARA'].str.strip().str.zfill(10)
    df['CAPITULO'] = pd.to_numeric(df['POSARA'].str[:2], errors='coerce')
    df['PARTE']    = pd.to_numeric(df['POSARA'].str[:4], errors='coerce')
    df['SA']       = df['POSARA'].str[:6]
    df['MES']      = pd.to_numeric(df['FECH'].str[2:4], errors='coerce')

    return df


def filtrar_impo(df: pd.DataFrame) -> pd.DataFrame:
    mask_reg = df['REGIMEN'].isin(REGIMEN_EXCLUIR)
    mask_air = (
        df['POSARA'].isin(POSARA_AEREAS)
        & df['REGIMEN'].isin(REGIMEN_AEREAS)
        & df['NIT'].isin(NITS_AEREAS)
    )
    df_filtrado = df[~(mask_reg | mask_air)].copy()
    n_excl = (mask_reg | mask_air).sum()
    print(f"  Excluidos: {n_excl:,}  |  Conservados: {len(df_filtrado):,}")
    return df_filtrado


def resumir(df: pd.DataFrame, sufijo: str) -> pd.DataFrame:
    freq = df.groupby(GRUPO, dropna=False).size().reset_index(name='_FREQ_')
    sumas = (
        df.groupby(GRUPO, as_index=False, dropna=False)[VARS_SUMA]
        .sum()
        .rename(columns={v: f'{v}{sufijo}' for v in VARS_SUMA})
    )
    return pd.merge(freq, sumas, on=GRUPO)


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("=== IMPORTACIONES NIT 2026 – MAÍZ / ARROZ / POLLO ===\n")

    print("Leyendo archivos 2024...")
    df2024 = filtrar_impo(leer_asu(ARCHIVOS_2024))
    print(f"  Total 2024: {len(df2024):,}\n")

    print("Leyendo archivos 2025...")
    df2025 = filtrar_impo(leer_asu(ARCHIVOS_2025))
    print(f"  Total 2025: {len(df2025):,}\n")

    print("Leyendo archivos 2026...")
    df2026 = filtrar_impo(leer_asu(ARCHIVOS_2026))
    print(f"  Total 2026: {len(df2026):,}\n")

    # PROC SUMMARY con filtro de POSARA dentro del macro
    tot24 = resumir(df2024, '2024')
    tot25 = resumir(df2025, '2025')
    tot26 = resumir(df2026, '2026')

    # DATA IMPO.TOTAL&VAR: MERGE TOTAL2026 TOTAL2025 TOTAL2024 (último gana en SAS)
    total = pd.merge(
        tot26.rename(columns={'_FREQ_': '_FREQ_2026'}),
        tot25.rename(columns={'_FREQ_': '_FREQ_2025'}),
        on=GRUPO, how='outer',
    )
    total = pd.merge(
        total,
        tot24.rename(columns={'_FREQ_': '_FREQ_2024'}),
        on=GRUPO, how='outer',
    )
    # _FREQ_: equivale a MERGE SAS donde tot24 es el último → 2024 gana sobre 2025 sobre 2026
    total['_FREQ_'] = (
        total['_FREQ_2024']
        .combine_first(total['_FREQ_2025'])
        .combine_first(total['_FREQ_2026'])
        .astype(int)
    )
    total = total.drop(columns=['_FREQ_2024', '_FREQ_2025', '_FREQ_2026'])
    # Posicionar _FREQ_ después de las claves de grupo
    resto = [c for c in total.columns if c not in GRUPO + ['_FREQ_']]
    total = total[GRUPO + ['_FREQ_'] + resto]
    total = total.sort_values('POSARA').reset_index(drop=True)

    # DATA EjeNIT: MERGE impo.totalposara(IN=A) IMPO.corre62; BY posara; IF A
    print("Leyendo tabla de referencia corre62...")
    # dtype=str preserva ceros iniciales ('012', '0015', etc.)
    # Solo strip en nombres: conserva capitalización original (Descrip, Cuode, CUCIsec…)
    corre62 = pd.read_excel(CORRE62_XLSX, dtype=str)
    corre62.columns = corre62.columns.str.strip()
    corre62['POSARA'] = corre62['POSARA'].str.strip().str.zfill(10)

    eje_nit = pd.merge(total, corre62, on='POSARA', how='left')

    # POSARA como entero, igual que SAS (elimina el cero de relleno a la izquierda)
    eje_nit['POSARA'] = pd.to_numeric(eje_nit['POSARA'], errors='coerce').astype('Int64')

    # IF A AND (CUCIsubg = '3341' OR CUCIsubg = '3343')
    # Con dtype=str, CUCIsubg es string → comparar con strings
    if 'CUCIsubg' in eje_nit.columns:
        eje_nit = eje_nit[eje_nit['CUCIsubg'].isin(['3341', '3343'])].copy()
    else:
        print("  [AVISO] Columna CUCIsubg no encontrada en corre62; se omite el filtro.")

    # Renombrar REGIMEN → regimen para coincidir con la variable SAS (definida en minúsculas)
    eje_nit = eje_nit.rename(columns={'REGIMEN': 'regimen'})

    print(f"Filas en EjeNIT: {len(eje_nit):,}")

    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        eje_nit.to_excel(writer, sheet_name='EJENIT', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
