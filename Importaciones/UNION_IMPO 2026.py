"""
Migración a Python: UNIÓN IMPO 2026
Equivalente al programa SAS 'UNIÓN_IMPO 2026.sas'.

Lee los archivos de importaciones de la Unión COMEX (formato 1071 chars),
corrige los códigos de país y genera un resumen CIF/Kilos netos por país.
Dependencias: pip install pandas openpyxl
"""

import io
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
    RUTA_FUENTE / 'M1ZF0226.txt',
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
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos → Python (pos-1, pos-1+ancho)
# Todos los campos del INPUT SAS incluidos
# =============================================================================

COLSPECS = [
    (0,    2),    # ANO        @1   $2.
    (2,    4),    # MES        @3   $2.
    (4,    6),    # ADUA       @5   2.
    (6,    9),    # PAISGEN    @7   $3.
    (9,    12),   # PAISPRO    @10  $3.
    (12,   15),   # PAISCOM    @13  $3.
    (15,   17),   # DEPTODES   @16  2.
    (17,   19),   # VIATRANS   @18  2.
    (19,   22),   # BANDERA    @20  $3.
    (22,   26),   # REGIMEN    @23  $4.
    (26,   29),   # ACUERDO    @27  $3.
    (29,   42),   # PBK        @30  13.2
    (42,   55),   # PNK        @43  13.2
    (55,   68),   # CANU       @56  13.2
    (68,   71),   # CODA       @69  $3.
    (71,   81),   # NABAN      @72  10.
    (81,   92),   # VAFODO     @82  11.2
    (92,   107),  # VAFOBP     @93  15.2
    (107,  118),  # FLETE      @108 11.2
    (118,  131),  # SEGUROS    @119 13.2
    (131,  144),  # OTROSG     @132 13.2
    (144,  157),  # VACIFD     @145 13.2
    (157,  172),  # VACIFP     @158 15.2
    (167,  177),  # UNICA      @168 10.   (el @SAS repositiona dentro de VACIFP)
    (177,  237),  # RIMPO      @178 $60.
    (237,  253),  # NIT        @238 $16.
    (253,  254),  # DIGV       @254 $1.
    (254,  257),  # CODONA     @255 3.
    (257,  258),  # OPER       @258 1.
    (258,  259),  # TIP_OPER   @259 1.
    (259,  260),  # TIP_USU    @260 1.
    (260,  263),  # COD_OPER   @261 $3.
    (263,  273),  # NUM_FOR    @264 10.
    (273,  274),  # impoexpo   @274 $1.
    (274,  275),  # tipodec    @275 $1.
    (275,  289),  # ndecant    @276 $14.
    (289,  297),  # fedecant   @290 8.
    (297,  305),  # nmanif     @298 $8.
    (305,  313),  # femanif    @306 8.
    (313,  329),  # ndoctran   @314 $16.
    (329,  331),  # CODADAD    @330 2.    (gap 331-336 no leído en SAS)
    (336,  341),  # cdeposit   @337 5.
    (341,  342),  # nitem      @342 1.
    (342,  402),  # decautor   @343 $60.
    (402,  403),  # tidendec   @403 $1.
    (403,  419),  # nidendec   @404 $16.
    (419,  420),  # DIGVFDEC   @420 $1.
    (420,  421),  # tidenimp   @421 $1.
    (421,  450),  # CIUDAIMP   @422 29.
    (450,  490),  # direcimp   @451 $40.
    (490,  550),  # exportad   @491 $60.
    (550,  551),  # tidenexp   @551 $1.
    (551,  567),  # nidenexp   @552 $16.
    (567,  568),  # digvfexp   @568 $1.
    (568,  603),  # direcexp   @569 $35.
    (603,  632),  # CIUDAEXP   @604 $29.
    (632,  636),  # ACTECON    @633 4.
    (636,  696),  # trportad   @637 $60.
    (696,  708),  # TIP_CAMB   @697 12.4
    (708,  710),  # ncuotas    @709 $2.
    (710,  721),  # VRAJUS     @711 11.2
    (721,  732),  # VADUA      @722 11.2
    (732,  736),  # cembal     @733 $4.
    (736,  745),  # nbultos    @737 9.
    (745,  747),  # subpart    @746 2.
    (747,  756),  # nlicenc    @748 $9.
    (756,  760),  # alicenc    @757 4.
    (760,  880),  # descmer    @761 $120.
    (880,  894),  # BASEIVA    @881 14.
    (894,  908),  # TOTALIVAYO @895 14.
    (908,  923),  # LUIN       @909 $15.
    (923,  926),  # CODLUIN    @924 $3.
    (926,  935),  # nlevante   @927 $9.
    (935,  943),  # felevant   @936 8.
    (943,  945),  # caduante   @944 2.
    (945,  959),  # numdecl    @946 $14.
    (959,  967),  # fedecla    @960 8.
    (967,  969),  # caduanex   @968 2.    (gap col 970 no leído en SAS)
    (970,  971),  # CLASE      @971 $1.
    (971,  973),  # DEPIM      @972 2.
    (973,  976),  # COPAEX     @974 $3.
    (976,  978),  # formapa    @977 2.
    (978,  980),  # TIPOIM     @979 2.
    (980,  988),  # PORARA     @981 8.2
    (988,  1002), # BASEARA    @989 14.
    (1002, 1016), # IMP1       @1003 14.
    (1016, 1024), # OTROSP     @1017 8.2
    (1024, 1038), # OTROSBASE  @1025 14.
    (1038, 1052), # OTDER      @1039 14.
    (1052, 1066), # DEREL      @1053 14.
    (1066, 1070), # ARTIC      @1067 4.
]

NOMBRES = [
    'ANO', 'MES', 'ADUA', 'PAISGEN', 'PAISPRO', 'PAISCOM',
    'DEPTODES', 'VIATRANS', 'BANDERA', 'REGIMEN', 'ACUERDO',
    'PBK', 'PNK', 'CANU', 'CODA', 'NABAN',
    'VAFODO', 'VAFOBP', 'FLETE', 'SEGUROS', 'OTROSG', 'VACIFD', 'VACIFP',
    'UNICA',
    'RIMPO', 'NIT', 'DIGV',
    'CODONA', 'OPER', 'TIP_OPER', 'TIP_USU', 'COD_OPER', 'NUM_FOR',
    'impoexpo', 'tipodec', 'ndecant', 'fedecant', 'nmanif', 'femanif',
    'ndoctran', 'CODADAD', 'cdeposit', 'nitem', 'decautor',
    'tidendec', 'nidendec', 'DIGVFDEC', 'tidenimp',
    'CIUDAIMP', 'direcimp', 'exportad',
    'tidenexp', 'nidenexp', 'digvfexp', 'direcexp', 'CIUDAEXP',
    'ACTECON', 'trportad', 'TIP_CAMB', 'ncuotas',
    'VRAJUS', 'VADUA', 'cembal', 'nbultos', 'subpart', 'nlicenc', 'alicenc',
    'descmer', 'BASEIVA', 'TOTALIVAYO',
    'LUIN', 'CODLUIN', 'nlevante', 'felevant', 'caduante',
    'numdecl', 'fedecla', 'caduanex',
    'CLASE', 'DEPIM', 'COPAEX', 'formapa', 'TIPOIM',
    'PORARA', 'BASEARA', 'IMP1', 'OTROSP', 'OTROSBASE', 'OTDER', 'DEREL', 'ARTIC',
]

COLUMNAS_TEXTO = {
    'ANO', 'MES', 'PAISGEN', 'PAISPRO', 'PAISCOM', 'BANDERA',
    'REGIMEN', 'ACUERDO', 'CODA', 'NABAN', 'RIMPO', 'NIT', 'DIGV',
    'COD_OPER', 'impoexpo', 'tipodec', 'ndecant', 'nmanif', 'ndoctran',
    'decautor', 'tidendec', 'nidendec', 'DIGVFDEC', 'tidenimp',
    'direcimp', 'exportad', 'tidenexp', 'nidenexp', 'digvfexp',
    'direcexp', 'CIUDAEXP', 'trportad', 'ncuotas', 'cembal', 'nlicenc',
    'descmer', 'LUIN', 'CODLUIN', 'nlevante', 'numdecl', 'CLASE', 'COPAEX',
}

# Campos con 2 decimales implícitos (informat w.2 en SAS: divide entre 100)
COLS_DEC2 = [
    'PBK', 'PNK', 'CANU', 'VAFODO', 'VAFOBP', 'FLETE',
    'SEGUROS', 'OTROSG', 'VACIFD', 'VACIFP',
    'VRAJUS', 'VADUA', 'PORARA', 'OTROSP',
]

# Campos con 4 decimales implícitos (informat w.4 en SAS: divide entre 10000)
COLS_DEC4 = ['TIP_CAMB']

# Códigos de país de 2 dígitos que necesitan cero a la izquierda
CODIGOS_RELLENAR = {
    '17', '23', '24', '26', '27', '29', '31', '32', '36', '40',
    '43', '48', '50', '53', '56', '59', '63', '68', '69', '72',
    '74', '76', '77', '80', '81', '83', '87', '90', '91', '93',
    '97', '98',
}

# =============================================================================
# FUNCIONES
# =============================================================================

def _leer_sin_bom(ruta) -> io.BytesIO:
    """Lee el archivo en binario y elimina el BOM de UTF-8 si está presente."""
    with open(ruta, 'rb') as f:
        datos = f.read()
    if datos[:3] == b'\xef\xbb\xbf':
        datos = datos[3:]
    return io.BytesIO(datos)


def _aplicar_decimal_implicito(serie: pd.Series, decimals: int = 2) -> pd.Series:
    """Replica el informat SAS w.d: divide entre 10**d si no hay punto en el dato."""
    divisor = 10 ** decimals
    def _conv(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        return float(s) if '.' in s else (float(s) / divisor if s else None)
    return serie.apply(_conv)


def leer_archivos(archivos: list) -> pd.DataFrame:
    """Equivale a DATA UNIMPO.UNION_IMPO2026; INFILE DATOIMPO — lee los 91 campos."""
    dtype = {col: str for col in COLUMNAS_TEXTO}
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            _leer_sin_bom(ruta),
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
        df[col] = _aplicar_decimal_implicito(df[col], decimals=2)

    for col in COLS_DEC4:
        df[col] = _aplicar_decimal_implicito(df[col], decimals=4)

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
    # salida_base  = RUTA_SALIDA / 'UNION_IMPO2026.xlsx'
    salida_cuadro = RUTA_SALIDA / 'UNION_IMPO_2026_CUADRO.xlsx'

    print(f"\nGuardando resultados en {RUTA_SALIDA}...")
    union_impo2026.to_excel(salida_base, index=False)
    cuadro.to_excel(salida_cuadro, index=False)

    print("Proceso finalizado correctamente.")
