"""
Migración a Python: UNIÓN IMPO 2026
Equivalente al programa SAS 'UNIÓN_IMPO 2026.sas'.

Lee los archivos de importaciones de la Unión COMEX (formato 1071 chars),
corrige los códigos de país y genera un resumen CIF/Kilos netos por país.
Dependencias: pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

RUTA_FUENTE = Path(r'\\systema75\migracion\UNION COMEX\UNION_COMEX_2026\UNION_IMPO')
RUTA_SALIDA = Path(r'D:\UNION\BASE_IMPO')

# Activar o comentar los meses disponibles
ARCHIVOS = [
    RUTA_FUENTE / 'M1ZF0126.txt',
    # RUTA_FUENTE / 'M1ZF0226.txt',
    # RUTA_FUENTE / 'M1ZF0326.txt',
    # RUTA_FUENTE / 'M1ZF0426.txt',
    # RUTA_FUENTE / 'M1ZF0526.txt',
    # RUTA_FUENTE / 'M1ZF0626.txt',
    # RUTA_FUENTE / 'M1ZF0726.txt',
    # RUTA_FUENTE / 'M1ZF0826.txt',
    # RUTA_FUENTE / 'M1ZF0926.txt',
    # RUTA_FUENTE / 'M1ZF1026.txt',
    # RUTA_FUENTE / 'M1ZF1126.txt',
    # RUTA_FUENTE / 'M1ZF1226.txt',
]

# =============================================================================
# LAYOUT DEL ARCHIVO DE ANCHO FIJO (LRECL=1071)
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos = Python pos-1
# Solo se leen los campos necesarios para el análisis
# =============================================================================

COLSPECS = [
    (0,   2),    # ANO       @1  $2.
    (2,   4),    # MES       @3  $2.
    (4,   6),    # ADUA      @5  2.
    (6,   9),    # PAISGEN   @7  $3.   (país generador/origen)
    (9,   12),   # PAISPRO   @10 $3.   (país productor)
    (12,  15),   # PAISCOM   @13 $3.   (país comercio)
    (15,  17),   # DEPTODES  @16 2.
    (17,  19),   # VIATRANS  @18 2.
    (19,  22),   # BANDERA   @20 $3.
    (22,  26),   # REGIMEN   @23 $4.
    (26,  29),   # ACUERDO   @27 $3.
    (29,  42),   # PBK       @30 13.2
    (42,  55),   # PNK       @43 13.2
    (55,  68),   # CANU      @56 13.2
    (68,  71),   # CODA      @69 $3.
    (71,  81),   # NABAN     @72 10.   (POSARA, 10 dígitos)
    (81,  92),   # VAFODO    @82 11.2  (valor FOB dólares)
    (92,  107),  # VAFOBP    @93 15.2  (valor FOB pesos)
    (107, 118),  # FLETE     @108 11.2
    (118, 131),  # SEGUROS   @119 13.2
    (131, 144),  # OTROSG    @132 13.2
    (144, 157),  # VACIFD    @145 13.2 (valor CIF dólares)
    (157, 172),  # VACIFP    @158 15.2 (valor CIF pesos)
    (177, 237),  # RIMPO     @178 $60.
    (237, 253),  # NIT       @238 $16.
    (253, 254),  # DIGV      @254 $1.
]

NOMBRES = [
    'ANO', 'MES', 'ADUA', 'PAISGEN', 'PAISPRO', 'PAISCOM',
    'DEPTODES', 'VIATRANS', 'BANDERA', 'REGIMEN', 'ACUERDO',
    'PBK', 'PNK', 'CANU', 'CODA', 'NABAN',
    'VAFODO', 'VAFOBP', 'FLETE', 'SEGUROS', 'OTROSG', 'VACIFD', 'VACIFP',
    'RIMPO', 'NIT', 'DIGV',
]

COLUMNAS_TEXTO = {
    'ANO', 'MES', 'PAISGEN', 'PAISPRO', 'PAISCOM', 'BANDERA',
    'REGIMEN', 'ACUERDO', 'CODA', 'NABAN', 'RIMPO', 'NIT', 'DIGV',
}

COLS_DEC2 = ['PBK', 'PNK', 'CANU', 'VAFODO', 'VAFOBP', 'FLETE', 'SEGUROS', 'OTROSG', 'VACIFD', 'VACIFP']

# Códigos de país de 2 dígitos que necesitan cero a la izquierda
# (Equivale al bloque IF PAISGEN = 'XX' THEN PAISGEN1 = '0XX' en SAS)
CODIGOS_RELLENAR = {
    '17', '23', '24', '26', '27', '29', '31', '32', '36', '40',
    '43', '48', '50', '53', '56', '59', '63', '68', '69', '72',
    '74', '76', '77', '80', '81', '83', '87', '90', '91', '93',
    '97', '98',
}

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


def leer_archivos(archivos: list) -> pd.DataFrame:
    """Equivale a DATA UNIMPO.UNION_IMPO2026; INFILE DATOIMPO."""
    dtype = {col: str for col in COLUMNAS_TEXTO}
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            ruta,
            colspecs=COLSPECS,
            names=NOMBRES,
            dtype=dtype,
            encoding='latin-1',
            header=None,
        )
        partes.append(df)
        print(f"  Leído: {Path(ruta).name}  ({len(df):,} registros)")

    df = pd.concat(partes, ignore_index=True)

    for col in NOMBRES:
        if col not in COLUMNAS_TEXTO:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for col in COLS_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    # Limpiar strings
    for col in COLUMNAS_TEXTO:
        df[col] = df[col].str.strip()

    df['NABAN'] = df['NABAN'].str.zfill(10)

    return df


def corregir_codigos_pais(df: pd.DataFrame) -> pd.DataFrame:
    """
    Equivale a DATA ARREGLO: normaliza PAISGEN a 3 dígitos con cero a la izquierda.
    AA = LENGTH(PAISGEN) para diagnóstico.
    """
    df = df.copy()
    df['PAISGEN1'] = df['PAISGEN']
    mask = df['PAISGEN1'].isin(CODIGOS_RELLENAR)
    df.loc[mask, 'PAISGEN1'] = df.loc[mask, 'PAISGEN1'].str.zfill(3)
    df['AA'] = df['PAISGEN1'].str.len()
    return df


def generar_cuadro(df: pd.DataFrame) -> pd.DataFrame:
    """
    Equivale a PROC SQL → CREATE TABLE CUADRO:
    SUM(VACIFD)/1000 y SUM(PNK)/1000 agrupados por PAISGEN1.
    """
    cuadro = (
        df.groupby('PAISGEN1', as_index=False)
        .agg(VACIFDT=('VACIFD', 'sum'), PNKT=('PNK', 'sum'))
    )
    cuadro['VACIFDT'] = cuadro['VACIFDT'] / 1000
    cuadro['PNKT']    = cuadro['PNKT']    / 1000
    return cuadro


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

if __name__ == '__main__':
    print("=== UNIÓN IMPO 2026 ===\n")

    print("Leyendo archivos fuente...")
    union_impo2026 = leer_archivos(ARCHIVOS)
    print(f"Total registros leídos: {len(union_impo2026):,}\n")

    print("Corrigiendo códigos de país...")
    arreglo = corregir_codigos_pais(union_impo2026)

    print("Generando cuadro resumen...")
    cuadro = generar_cuadro(arreglo)
    print(f"Países distintos: {len(cuadro)}\n")

    print("Cuadro resumen (VACIFD y PNK en miles de dólares):")
    pd.set_option('display.float_format', '{:,.2f}'.format)
    print(cuadro.to_string(index=False))

    # Guardar resultados
    RUTA_SALIDA.mkdir(parents=True, exist_ok=True)
    salida_excel = RUTA_SALIDA / 'UNION_IMPO_2026_CUADRO.xlsx'
    salida_base  = RUTA_SALIDA / 'UNION_IMPO_2026_BASE.csv'

    print(f"\nGuardando resultados en {RUTA_SALIDA}...")
    cuadro.to_excel(salida_excel, index=False)
    arreglo.to_csv(salida_base, index=False, encoding='latin-1')

    print("Proceso finalizado correctamente.")
