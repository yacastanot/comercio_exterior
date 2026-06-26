"""
Migración a Python: Importaciones 2026 (enero-febrero)
Equivalente al programa SAS 'Impo 2025 a 2026_Total_Ultima versión.sas'.

Procesa solo los meses disponibles de 2026.
Agrupación: posara, parte, capitulo, depto, paor, mes.
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
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Importaciones 2026.xlsx'
CORRE62_XLSX   = RUTA_COMEX / 'corre62.xlsx'

# # 2024: año completo (12 meses)
# ARCHIVOS_2024 = [
#     RUTA_ASU / 'M10124.asu', RUTA_ASU / 'M10224.asu',
#     RUTA_ASU / 'M10324.asu', RUTA_ASU / 'M10424.asu',
#     RUTA_ASU / 'M10524.asu', RUTA_ASU / 'M10624.asu',
#     RUTA_ASU / 'M10724.asu', RUTA_ASU / 'M10824.asu',
#     RUTA_ASU / 'M10924.asu', RUTA_ASU / 'M11024.asu',
#     RUTA_ASU / 'M11124.asu', RUTA_ASU / 'M11224.asu',
# ]

# 2025: mismo periodo que 2026 (enero-febrero)
ARCHIVOS_2025 = [
    RUTA_ASU / 'M10125.asu',
    RUTA_ASU / 'M10225.asu',
    # RUTA_ASU / 'M10325.asu',
    # RUTA_ASU / 'M10425.asu',
    # RUTA_ASU / 'M10525.asu',
    # RUTA_ASU / 'M10625.asu',
    # RUTA_ASU / 'M10725.asu',
    # RUTA_ASU / 'M10825.asu',
    # RUTA_ASU / 'M10925.asu',
    # RUTA_ASU / 'M11025.asu',
    # RUTA_ASU / 'M11125.asu',
    # RUTA_ASU / 'M11225.asu',
]

# 2026: meses disponibles (comentar los no disponibles)
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
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos = Python pos-1
# =============================================================================

COLSPECS = [
    (0,  4),    # FECH     @1  $4.   (YYMM)
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
    'FECH', 'PAOR', 'PAPR', 'PACO', 'REGIMEN', 'ACUERDO', 'CODUN',
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
GRUPO     = ['POSARA', 'PARTE', 'CAPITULO', 'DEPTO', 'PAOR', 'MES']
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
    """Lee y concatena archivos .asu de ancho fijo."""
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
    print("=== IMPORTACIONES 2026 vs 2025 (enero-febrero) ===\n")

    # print("Leyendo archivos 2024...")
    # df2024 = filtrar_impo(leer_asu(ARCHIVOS_2024))
    # print(f"  Total 2024: {len(df2024):,}\n")

    print("Leyendo archivos 2025...")
    df2025 = filtrar_impo(leer_asu(ARCHIVOS_2025))
    print(f"  Total 2025: {len(df2025):,}\n")

    print("Leyendo archivos 2026...")
    df2026 = filtrar_impo(leer_asu(ARCHIVOS_2026))
    print(f"  Total 2026: {len(df2026):,}\n")

    # Equivale a PROC SUMMARY + DATA IMPO.TOTAL&VAR (merge outer, 2026 primero)
    # tot24 = resumir(df2024, '2024')
    tot25 = resumir(df2025, '2025')
    tot26 = resumir(df2026, '2026')
    total = pd.merge(
        tot26,
        tot25.rename(columns={'_FREQ_': '_FREQ_2025'}),
        on=GRUPO, how='outer',
    )
    # _FREQ_ toma el valor de 2026 si existe, si no el de 2025
    total['_FREQ_'] = total['_FREQ_'].combine_first(total.pop('_FREQ_2025')).astype(int)
    # total = pd.merge(total, tot24, on=GRUPO, how='outer')
    total = total.sort_values('POSARA').reset_index(drop=True)

    print("Leyendo tabla de referencia corre62...")
    # dtype=str preserva ceros iniciales ('012', '0015', etc.)
    # Solo strip en nombres: conserva capitalización original (Descrip, Cuode, CUCIsec…)
    corre62 = pd.read_excel(CORRE62_XLSX, dtype=str)
    corre62.columns = corre62.columns.str.strip()
    corre62['POSARA'] = corre62['POSARA'].str.strip().str.zfill(10)

    eje = pd.merge(total, corre62, on='POSARA', how='left')
    print(f"Filas en EJE: {len(eje):,}")

    # POSARA como entero, igual que SAS (elimina el cero de relleno a la izquierda)
    eje['POSARA'] = pd.to_numeric(eje['POSARA'], errors='coerce').astype('Int64')

    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    MAX_FILAS = 1_048_575  # límite de filas de Excel menos 1 para el encabezado
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        if len(eje) <= MAX_FILAS:
            eje.to_excel(writer, sheet_name='EJE', index=False)
        else:
            for i, inicio in enumerate(range(0, len(eje), MAX_FILAS), start=1):
                eje.iloc[inicio:inicio + MAX_FILAS].to_excel(
                    writer, sheet_name=f'EJE_{i}', index=False
                )
            print(f"  Datos divididos en {i} hojas (EJE_1 … EJE_{i})")

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
