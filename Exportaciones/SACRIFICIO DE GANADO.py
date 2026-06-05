"""
Migración a Python: SACRIFICIO DE GANADO
Equivalente al programa SAS original del mismo nombre.

Combina exportaciones (.AVA) e importaciones (.ASU) de los últimos meses del año,
filtra posiciones arancelarias de ganado bovino, porcino y ovino/caprino,
cruza con tabla de países y genera resumen por múltiples variables.

Dependencias:
    pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path
from datetime import date

# =============================================================================
# CONFIGURACIÓN DE RUTAS Y AÑO
# =============================================================================

ANO = 25   # Equivale a %EXPO25(25) / %IMPO25(25) en el SAS

RUTA_EXPO   = Path(r'\\systema44\Migracion\E110')
RUTA_IMPO   = Path(r'\\systema44\Migracion\E100')
RUTA_PAIS   = Path(r'\\systema44\Migracion\E125') / f'PA20{ANO}'
RUTA_SALIDA = Path(r'D:\Migracion\SACRIFICIO DE GANADO')
#RUTA_SALIDA = Path(r'\\systema75\Migracion\SACRIFICIO DE GANADO')
ARCHIVO_SALIDA = RUTA_SALIDA / f'SACRIFICIO DE GANADO {date.today():%d%B%Y}.xlsx'

# Archivos de exportaciones: meses 10, 11, 12 del año
ARCHIVOS_EXPO = [
    RUTA_EXPO / f'M110{ANO}.AVA',
    RUTA_EXPO / f'M111{ANO}.AVA',
    RUTA_EXPO / f'M112{ANO}.AVA',
]

# Archivos de importaciones: meses 09, 10, 11 del año
ARCHIVOS_IMPO = [
    RUTA_IMPO / f'M109{ANO}.ASU',
    RUTA_IMPO / f'M110{ANO}.ASU',
    RUTA_IMPO / f'M111{ANO}.ASU',
]

# =============================================================================
# DEFINICIÓN EXPORTACIONES (LRECL=406)
# SAS usa notación de columnas (col_ini-col_fin, 1-indexed, inclusive)
# Python: (col_ini - 1, col_fin)
# =============================================================================

COLSPECS_EXPO = [
    (2,   4),    # ANS      $3-4
    (4,   6),    # MES      $5-6
    (15,  27),   # NIT      $16-27
    (41,  44),   # PAIS     $42-44
    (71,  75),   # CAPI4    $153→72-75? — ver nota abajo
    (143, 146),  # CODMODA  $144-146
    (146, 147),  # PAGO     147
    (152, 156),  # CAPI4    $153-156
    (152, 158),  # CAPI6    $153-158
    (152, 162),  # NANDINA  153-162
    (177, 192),  # CANTI    178-192 .2
    (192, 207),  # PBK      193-207 .2
    (207, 222),  # PNK      208-222 .2
    (222, 237),  # FOBDO    223-237 .2
    (346, 406),  # RAZON    $347-406
]

NOMBRES_EXPO = [
    'ANS', 'MES', 'NIT', 'PAIS',
    '_CAPI4_TMP',   # columna auxiliar (no usada, ver abajo)
    'CODMODA', 'PAGO',
    'CAPI4', 'CAPI6', 'NANDINA',
    'CANTI', 'PBK', 'PNK', 'FOBDO', 'RAZON',
]

# =============================================================================
# DEFINICIÓN IMPORTACIONES (LRECL=440)
# =============================================================================

COLSPECS_IMPO = [
    (0,   2),    # ANS      $1-2
    (2,   4),    # MES      $3-4
    (6,   9),    # PAIS     $7-9
    (22,  26),   # REGIMEN  $23-26
    (29,  42),   # PBK      30-42 .2
    (42,  55),   # PNK      43-55 .2
    (55,  68),   # CANTI    @56 13.2  (posición @56, ancho 13)
    (71,  75),   # CAPI4    $72-75
    (71,  77),   # CAPI6    $72-77
    (71,  81),   # NANDINA  72-81
    (89,  100),  # FOBDO    90-100 .2
    (111, 124),  # CIFDOL   112-124 .2
    (363, 379),  # NIT      $364-379
    (380, 440),  # RAZON    $381-440
]

NOMBRES_IMPO = [
    'ANS', 'MES', 'PAIS', 'REGIMEN',
    'PBK', 'PNK', 'CANTI',
    'CAPI4', 'CAPI6', 'NANDINA',
    'FOBDO', 'CIFDOL',
    'NIT', 'RAZON',
]

# Columnas con 2 decimales implícitos (SAS w.2 / col .2)
COLS_DEC2_EXPO = ['CANTI', 'PBK', 'PNK', 'FOBDO']
COLS_DEC2_IMPO = ['PBK', 'PNK', 'CANTI', 'FOBDO', 'CIFDOL']

# Capítulos arancelarios a conservar (bovino, porcino, ovino/caprino)
CAPI4_GANADO = {'0102', '0201', '0202', '0103', '0203', '0204'}

# Variables numéricas a sumar en el resumen
VARS_SUMA = ['FOBDO', 'CIFDOL', 'CANTI', 'PNK']

# Variables de agrupación del macro %SUMMARY
GRUPOS = ['TIPO', 'ANS', 'MES', 'CAPI4', 'CAPI6', 'NANDINA',
          'PAISORI_DES', 'PAISDES', 'NIT', 'RAZON']

# =============================================================================
# FILTROS EXPORTACIONES
# =============================================================================

CODMODA_EXCLUIR = {
    '001', '003', '105', '301', '302', '303', '304', '305',
    '306', '307', '399', '308', '309', '405', '701', '500', '202',
}
CODMODA_FORPAGO2  = {'401', '402', '403'}
NANDINA_EXCLUIR   = {'7108200000', '4907002000', '4907003000'}

# =============================================================================
# FILTROS IMPORTACIONES
# =============================================================================

REGIMEN_EXCLUIR = {
    'C390', 'C392', 'C393', 'C394', 'C395', 'C396', 'C397',
    'C540', 'C541', 'C545', 'C546', 'C660', 'C665', 'S100',
    'S105', 'S106', 'S200', 'S140', 'S240', 'C398', 'C547',
}
NANDINA_EXCLUIR_IMPO = {'9803000000'}

# Aerolíneas (aeronaves que no son importaciones comerciales reales)
AERONAVES_IMPO = [
    {'nit': {'0000000890912462'},        'regimen': {'C196'}, 'nandina': {'8802400000'}},
    {'nit': {'00000890100577', '000890100577', '0000000890100577'},
     'regimen': {'C190'}, 'nandina': {'8802400000'}},
]

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    """Replica el informat SAS w.2: divide entre 100 si no hay punto en el dato."""
    def _conv(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        return float(s) if '.' in s else (float(s) / 100 if s else None)
    return serie.apply(_conv)


def _leer_fwf(archivos: list, colspecs: list, nombres: list,
              cols_str: set, cols_dec2: list) -> pd.DataFrame:
    """Lee y concatena archivos de ancho fijo."""
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            ruta,
            colspecs=colspecs,
            names=nombres,
            dtype={c: str for c in cols_str},
            header=None,
            encoding='latin-1',
        )
        partes.append(df)

    df = pd.concat(partes, ignore_index=True)

    for col in nombres:
        if col not in cols_str:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    for col in cols_dec2:
        df[col] = _aplicar_decimal_implicito(df[col])

    return df


def leer_exportaciones(archivos: list) -> pd.DataFrame:
    """Macro %EXPO25: lee .AVA, aplica filtros y etiqueta como EXPO."""
    cols_str = {'ANS', 'MES', 'NIT', 'PAIS', '_CAPI4_TMP',
                'CODMODA', 'CAPI4', 'CAPI6', 'NANDINA', 'RAZON'}

    df = _leer_fwf(archivos, COLSPECS_EXPO, NOMBRES_EXPO, cols_str, COLS_DEC2_EXPO)
    df.drop(columns=['_CAPI4_TMP'], inplace=True)

    # Normalizar
    df['NANDINA']  = df['NANDINA'].str.strip().str.zfill(10)
    df['CODMODA']  = df['CODMODA'].str.strip()
    df['CAPI4']    = df['CAPI4'].str.strip()
    df['CAPI6']    = df['CAPI6'].str.strip()
    df['NIT']      = df['NIT'].str.strip()
    df['RAZON']    = df['RAZON'].str.strip()
    df['CIFDOL']   = float('nan')   # Campo exclusivo de importaciones
    df['TIPO']     = 'EXPO'

    # Filtros de exclusión (IF ... THEN DELETE)
    mask1 = df['CODMODA'].isin(CODMODA_EXCLUIR)
    mask2 = df['CODMODA'].isin(CODMODA_FORPAGO2) & (df['PAGO'] == 2)
    mask3 = df['NANDINA'].isin(NANDINA_EXCLUIR)
    df = df[~(mask1 | mask2 | mask3)].copy()

    n_excluidos = (mask1 | mask2 | mask3).sum()
    print(f"  Exportaciones — leídos: {len(df) + n_excluidos:,}  |  excluidos: {n_excluidos:,}  |  conservados: {len(df):,}")
    return df


def leer_importaciones(archivos: list) -> pd.DataFrame:
    """Macro %IMPO25: lee .ASU, aplica filtros y etiqueta como IMPO."""
    cols_str = {'ANS', 'MES', 'PAIS', 'REGIMEN',
                'CAPI4', 'CAPI6', 'NANDINA', 'NIT', 'RAZON'}

    df = _leer_fwf(archivos, COLSPECS_IMPO, NOMBRES_IMPO, cols_str, COLS_DEC2_IMPO)

    # Normalizar
    df['NANDINA']  = df['NANDINA'].str.strip().str.zfill(10)
    df['REGIMEN']  = df['REGIMEN'].str.strip()
    df['CAPI4']    = df['CAPI4'].str.strip()
    df['CAPI6']    = df['CAPI6'].str.strip()
    df['NIT']      = df['NIT'].str.strip()
    df['RAZON']    = df['RAZON'].str.strip()
    df['CODMODA']  = float('nan')   # Campo exclusivo de exportaciones
    df['PAGO']     = float('nan')
    df['PBK']      = df.get('PBK', float('nan'))
    df['TIPO']     = 'IMPO'

    # Filtros de exclusión
    mask1 = df['REGIMEN'].isin(REGIMEN_EXCLUIR) | df['NANDINA'].isin(NANDINA_EXCLUIR_IMPO)
    mask2 = pd.Series(False, index=df.index)
    for regla in AERONAVES_IMPO:
        mask2 |= (
            df['NIT'].isin(regla['nit'])
            & df['REGIMEN'].isin(regla['regimen'])
            & df['NANDINA'].isin(regla['nandina'])
        )
    df = df[~(mask1 | mask2)].copy()

    n_excluidos = (mask1 | mask2).sum()
    print(f"  Importaciones  — leídas: {len(df) + n_excluidos:,}  |  excluidas: {n_excluidos:,}  |  conservadas: {len(df):,}")
    return df


def normalizar_pais(df: pd.DataFrame) -> pd.DataFrame:
    """
    Equivale al segundo DATA COMERCIO del SAS:
    rellena PAIS con cero a la izquierda si tiene menos de 3 caracteres.
    """
    df = df.copy()
    df['PAIS'] = df['PAIS'].fillna('').str.strip()
    df['PAISORI_DES'] = df['PAIS'].str.zfill(3)
    return df


def leer_tabla_paises(ruta: Path) -> pd.DataFrame:
    """
    Lee el archivo de referencia de países PA20XX (ancho fijo, sin extensión):
    PAISORI_DES $9-11  PAISDES $14-43

    Usa corte por bytes (igual que SAS) para que nombres con caracteres
    multibyte en UTF-8 queden truncados en la misma posición que en SAS.
    """
    registros = []
    raw = Path(ruta).read_bytes()
    for linea in raw.split(b'\n'):
        linea = linea.rstrip(b'\r')
        if len(linea) < 11:
            continue
        paisori = linea[8:11].decode('utf-8', errors='replace').strip().zfill(3)
        paisdes = linea[13:43].decode('utf-8', errors='replace').strip() if len(linea) > 13 else ''
        registros.append({'PAISORI_DES': paisori, 'PAISDES': paisdes})
    df = pd.DataFrame(registros)
    return df[df['PAISORI_DES'].str.strip() != ''].reset_index(drop=True)


def resumir(df: pd.DataFrame, grupos: list, vars_suma: list) -> pd.DataFrame:
    """
    Equivale al macro %SUMMARY:
    agrega sumando vars_suma agrupando por grupos.
    """
    return df.groupby(grupos, as_index=False, dropna=False)[vars_suma].sum()


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    # --- Leer exportaciones e importaciones ---
    print(f"Leyendo exportaciones (año 20{ANO}, meses 10-12)...")
    expo = leer_exportaciones(ARCHIVOS_EXPO)

    print(f"Leyendo importaciones (año 20{ANO}, meses 09-11)...")
    impo = leer_importaciones(ARCHIVOS_IMPO)

    # --- Combinar y filtrar por capítulo arancelario ---
    # DATA COMERCIO; SET IMP25 EXP25; IF CAPI4 IN(...)
    comercio = pd.concat([impo, expo], ignore_index=True)
    comercio = comercio[comercio['CAPI4'].isin(CAPI4_GANADO)].copy()
    print(f"\nRegistros ganado/carne (CAPI4 filtrado): {len(comercio):,}")

    # --- Normalizar código de país ---
    comercio = normalizar_pais(comercio)

    # --- Leer tabla de países y hacer merge ---
    print(f"Leyendo tabla de países: {RUTA_PAIS}")
    paises = leer_tabla_paises(RUTA_PAIS)

    # MERGE COMERCIO(IN=N) PAIS; BY PAISORI_DES; IF N  →  LEFT JOIN
    comercio = pd.merge(comercio, paises, on='PAISORI_DES', how='left')
    print(f"Registros tras cruce con países: {len(comercio):,}")

    # --- Resumen (%SUMMARY → RESUL1) ---
    print("Generando resumen RESUL1...")
    resul1 = resumir(comercio, GRUPOS, VARS_SUMA)
    print(f"Filas en RESUL1: {len(resul1):,}")

    # --- Exportar a Excel ---
    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        resul1.to_excel(writer, sheet_name='RESUL1', index=False)
        comercio.to_excel(writer, sheet_name='COMERCIO', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
