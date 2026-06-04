"""
Migración a Python: Expo 2025-2026 Economía Naranja
Equivalente al programa SAS original del mismo nombre.

Dependencias:
    pip install pandas pyreadstat openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

RUTA_AVA       = Path(r'\\systema44\Migracion\e110')
RUTA_COMEX     = Path(r'D:\COMEX')
RUTA_RESULTADO = RUTA_COMEX / 'EXPO' / 'Resultados'
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Exportaciones_EN_2026.xlsx'
CORRE62_XLSX   = RUTA_COMEX / 'corre62.xlsx'

ARCHIVOS_2025 = [
    RUTA_AVA / 'M10125.ava',
    RUTA_AVA / 'M10225.ava',
    # RUTA_AVA / 'M10325.ava',
    # RUTA_AVA / 'M10425.ava',
    # RUTA_AVA / 'M10525.ava',
    # RUTA_AVA / 'M10625.ava',
    # RUTA_AVA / 'M10725.ava',
    # RUTA_AVA / 'M10825.ava',
    # RUTA_AVA / 'M10925.ava',
    # RUTA_AVA / 'M11025.ava',
    # RUTA_AVA / 'M11125.ava',
    # RUTA_AVA / 'M11225.ava',
]

ARCHIVOS_2026 = [
    RUTA_AVA / 'M10126.ava',
    RUTA_AVA / 'M10226.ava',
    # RUTA_AVA / 'M10326.ava',
    # RUTA_AVA / 'M10426.ava',
    # RUTA_AVA / 'M10526.ava',
    # RUTA_AVA / 'M10626.ava',
    # RUTA_AVA / 'M10726.ava',
    # RUTA_AVA / 'M10826.ava',
    # RUTA_AVA / 'M10926.ava',
    # RUTA_AVA / 'M11026.ava',
    # RUTA_AVA / 'M11126.ava',
    # RUTA_AVA / 'M11226.ava',
]

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO (LRECL=678)
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos = Python pos-1
# =============================================================================

COLSPECS = [
    (0,   4),    # FECHPRO   @1  4.
    (4,   6),    # MES       @5  2.
    (12,  14),   # ADU       @13 2.
    (15,  27),   # NIT       @16 $12.
    (36,  41),   # MUNICIPIO @37 5.
    (41,  44),   # PAIS      @42 3.
    (47,  67),   # CIUEXP    @48 $20.
    (84,  86),   # LUGSALN   @85 2.
    (86,  89),   # LUGSAL    @87 $3.
    (89,  91),   # DEPTPR    @90 2.
    (91,  105),  # DEXANTE   @92 14.   (DEX ANTE en el SAS original)
    (138, 139),  # VIA       @139 1.
    (139, 142),  # BANDERA   @140 3.
    (142, 143),  # REGIMEN   @143 1.
    (143, 146),  # MODALIDAD @144 3.
    (146, 147),  # FORPA4    @147 1.
    (152, 162),  # POSARA    @153 10.  (leído como string para preservar ceros)
    (170, 172),  # DEPTOR    @171 2.
    (172, 174),  # UNIDAD    @173 2.
    (177, 192),  # CANTIDAD  @178 15.2
    (207, 222),  # KNETO     @208 15.2
    (222, 237),  # FOBDOL    @223 15.2
    (237, 257),  # FOBPES    @238 20.2
    (272, 287),  # FLETES    @273 15.
    (287, 302),  # SEGUROS   @288 15.
    (319, 325),  # FECHEMB   @320 6.
    (325, 338),  # DEX       @326 $13.
    (338, 346),  # FECHDEC   @339 8.
    (346, 406),  # RAZON     @347 $60.
]

NOMBRES = [
    'FECHPRO', 'MES', 'ADU', 'NIT', 'MUNICIPIO', 'PAIS', 'CIUEXP',
    'LUGSALN', 'LUGSAL', 'DEPTPR', 'DEXANTE', 'VIA', 'BANDERA',
    'REGIMEN', 'MODALIDAD', 'FORPA4', 'POSARA', 'DEPTOR', 'UNIDAD',
    'CANTIDAD', 'KNETO', 'FOBDOL', 'FOBPES', 'FLETES', 'SEGUROS',
    'FECHEMB', 'DEX', 'FECHDEC', 'RAZON',
]

# Columnas que se leen como texto
COLS_STR = {
    'NIT', 'CIUEXP', 'LUGSAL', 'POSARA', 'DEX', 'FECHDEC', 'FECHEMB',
    'FECHPRO', 'RAZON',
}

# Columnas con 2 decimales implícitos (SAS w.2): si el valor no contiene
# punto decimal en el archivo, se divide entre 100.
COLS_IMPLIED_DEC2 = ['CANTIDAD', 'KNETO', 'FOBDOL', 'FOBPES']

# =============================================================================
# POSICIONES ARANCELARIAS DE ECONOMÍA NARANJA
# =============================================================================

POSARA_EN = {
    '3704000000', '3705000000', '3705100000', '3705100010', '3705100090',
    '3705200000', '3705200010', '3705200090', '3705900000', '3705900010',
    '3705900090', '3706100000', '3706100010', '3706100011', '3706100012',
    '3706100021', '3706100022', '3706100023', '3706100024', '3706100025',
    '3706100090', '3706100091', '3706100092', '3706100093', '3706100094',
    '3706100095', '3706900000', '3706900010', '3706900011', '3706900012',
    '3706900021', '3706900022', '3706900023', '3706900024', '3706900025',
    '3706900090', '3706900091', '3706900092', '3706900093', '3706900094',
    '3706900095',
    '4901100010', '4901100020', '4901100090', '4901101000', '4901109000',
    '4901910000', '4901990010', '4901990020', '4901990090', '4901991000',
    '4901999000', '4902100000', '4902900010', '4902900090', '4902901000',
    '4902909000', '4903000000', '4904000000', '4905200000', '4905910000',
    '4905990000', '4908100000', '4908900010', '4908900090', '4908901000',
    '4908909000', '4909000000', '4910000000', '4911100000', '4911101000',
    '4911109000', '4911910000', '4911911000', '4911919000',
    '8523402900', '8523491000', '8523492000', '8523499000',
    '8524101000', '8524109000', '8524211000', '8524212000', '8524219000',
    '8524221000', '8524222000', '8524229000', '8524229010', '8524229090',
    '8524231000', '8524232000', '8524239000', '8524239010', '8524239090',
    '8524310000', '8524320000', '8524390000', '8524400000',
    '8524511000', '8524519000', '8524521000', '8524529000',
    '8524531000', '8524539000', '8524600000', '8524901000', '8524902000',
    '8524909010', '8524909090', '8524909100', '8524909900',
    '8524991000', '8524999000',
    '9201100000', '9201200000', '9201900000', '9202100000', '9202900000',
    '9203000000', '9204100000', '9204200000', '9205100000', '9205900000',
    '9205901000', '9205902000', '9205903000', '9205909000', '9206000000',
    '9207100000', '9207900000', '9208100000', '9208900000',
    '9209100000', '9209200000', '9209300000', '9209910000', '9209920000',
    '9209930000', '9209940000', '9209990000',
    '9501000000', '9502100000', '9502910000', '9502990000',
    '9503001000', '9503002000', '9503002100', '9503002200', '9503002800',
    '9503002900', '9503002920', '9503002990', '9503003000', '9503004000',
    '9503009100', '9503009200', '9503009300', '9503009400', '9503009500',
    '9503009600', '9503009900', '9503009910', '9503009990',
    '9503100000', '9503200000', '9503200010', '9503200090',
    '9503300000', '9503300010', '9503300090',
    '9503410000', '9503490000', '9503490010', '9503490090',
    '9503500000', '9503500010', '9503500090',
    '9503600000', '9503600010', '9503600090',
    '9503700000', '9503700010', '9503700090',
    '9503800000', '9503900000',
    '9504200000', '9504300000', '9504301000', '9504301010', '9504301090',
    '9504309000', '9504400000', '9504901000', '9504902000', '9504903000',
    '9504909000', '9504909100', '9504909900',
    '9701100000', '9701100010', '9701100090', '9701210000', '9701220000',
    '9701290000', '9701900000', '9701900010', '9701900090', '9701910000',
    '9701920000', '9701990000', '9702000000', '9702000010', '9702000090',
    '9702100000', '9702900000', '9703000000', '9703000010', '9703000090',
    '9703100000', '9703900000', '9705000000', '9705000010', '9705000090',
    '9705100000', '9705210000', '9705220000', '9705290000', '9705310000',
    '9705390000', '9706000000', '9706000010', '9706000090', '9706100000',
    '9706900000',
    '9804000000', '9806000000', '9809000000',
}

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    """
    Replica el informat SAS w.2:
    si el valor en el archivo no contiene punto decimal, divide entre 100.
    """
    def _convertir(val):
        if pd.isna(val):
            return val
        s = str(val).strip()
        if not s:
            return None
        return float(s) if '.' in s else float(s) / 100

    return serie.apply(_convertir)


def leer_ava(archivos: list) -> pd.DataFrame:
    """Lee y concatena archivos .ava de ancho fijo (equivale a DATA + INFILE)."""
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            ruta,
            colspecs=COLSPECS,
            names=NOMBRES,
            dtype={col: str for col in COLS_STR},
            header=None,
            encoding='latin-1',
        )
        partes.append(df)

    df = pd.concat(partes, ignore_index=True)

    # Convertir columnas numéricas
    for col in NOMBRES:
        if col not in COLS_STR:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Decimales implícitos (SAS w.2)
    for col in COLS_IMPLIED_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    # Normalizar POSARA a string de 10 dígitos con ceros a la izquierda
    df['POSARA'] = df['POSARA'].str.strip().str.zfill(10)

    # Derivar campos arancelarios desde POSARA
    # SAS: @153 CAPITULO 2. / PARTE 4. / SA 6. y luego parte=substr(posara,1,6)
    df['CAPITULO'] = df['POSARA'].str[:2].astype(int)
    df['PARTE']    = df['POSARA'].str[:6]   # substr(posara,1,6) del SAS
    df['SA']       = df['POSARA'].str[:6]

    # Indicador TRA: excluye productos minero-energéticos
    # TRA=1 → petróleo/carbón/café  |  TRA=2 → el resto
    posara_tra = {'0901110000', '0901111000', '0901120000', '7202600000'}
    parte_int  = df['POSARA'].str[:4].astype(int)

    mask_tra = (
        df['POSARA'].isin(posara_tra)
        | parte_int.between(2701, 2704)
        | parte_int.between(2709, 2715)
    )
    df['TRA'] = 2
    df.loc[mask_tra, 'TRA'] = 1

    return df


def calc(df2025: pd.DataFrame, df2026: pd.DataFrame, grupo: str | list) -> pd.DataFrame:
    """
    Equivale al macro SAS %CALC:
    suma FOBDOL, KNETO, FOBPES y CANTIDAD para cada año y hace merge externo.
    Incluye _FREQ_ replicando el comportamiento de PROC SUMMARY: usa el conteo
    de 2026 cuando existe, sino el de 2025 (igual que el DATA MERGE de SAS).
    """
    if isinstance(grupo, str):
        grupo = [grupo]

    vars_suma = ['FOBDOL', 'KNETO', 'FOBPES', 'CANTIDAD']

    cnt25 = df2025.groupby(grupo).size().reset_index(name='_FREQ_25')
    cnt26 = df2026.groupby(grupo).size().reset_index(name='_FREQ_26')

    tot25 = (
        df2025.groupby(grupo, as_index=False)[vars_suma]
        .sum()
        .rename(columns={v: f'{v}2025' for v in vars_suma})
        .merge(cnt25, on=grupo)
    )
    tot26 = (
        df2026.groupby(grupo, as_index=False)[vars_suma]
        .sum()
        .rename(columns={v: f'{v}2026' for v in vars_suma})
        .merge(cnt26, on=grupo)
    )

    merged = pd.merge(tot25, tot26, on=grupo, how='outer')
    merged['_FREQ_'] = merged['_FREQ_26'].combine_first(merged['_FREQ_25']).astype(int)
    merged = merged.drop(columns=['_FREQ_25', '_FREQ_26'])
    return merged


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("Leyendo archivos 2025...")
    df2025 = leer_ava(ARCHIVOS_2025)
    print(f"  Registros 2025: {len(df2025):,}")

    print("Leyendo archivos 2026...")
    df2026 = leer_ava(ARCHIVOS_2026)
    print(f"  Registros 2026: {len(df2026):,}")

    # Equivale a %CALC(posara parte capitulo deptor pais mes):
    # agrupación conjunta por los 6 campos, luego filtro POSARA_EN
    GRUPO_EJE = ['POSARA', 'PARTE', 'CAPITULO', 'DEPTOR', 'PAIS', 'MES']

    print(f"  Calculando EXPO.TOTALposara parte capitulo deptor pais mes...")
    total_detalle = calc(df2025, df2026, GRUPO_EJE)

    # Equivale a: IF POSARA IN (...) dentro del DATA EXPO.TOTAL&VAR
    total_detalle = total_detalle[
        total_detalle['POSARA'].isin(POSARA_EN)
    ].copy()

    # Equivale a: PROC SORT DATA=EXPO.totalposara; BY posara;
    total_detalle = total_detalle.sort_values('POSARA').reset_index(drop=True)

    # Equivale a: DATA expo.EJE; MERGE expo.totalposara(IN=A) EXPO.corre62; BY posara; IF A;
    print("Leyendo tabla de referencia corre62...")
    corre62 = pd.read_excel(CORRE62_XLSX)
    corre62.columns = corre62.columns.str.upper()
    corre62['POSARA'] = corre62['POSARA'].astype(str).str.strip().str.zfill(10)

    eje = pd.merge(total_detalle, corre62, on='POSARA', how='left')

    # Renombrar columnas para coincidir exactamente con la salida SAS
    eje = eje.rename(columns={
        'POSARA':     'posara',
        'DEPTOR':     'deptor',
        'PAIS':       'pais',
        'FOBPES2025': 'fobpes2025',
        'CANTIDAD2025':'cantidad2025',
        'FOBPES2026': 'fobpes2026',
        'CANTIDAD2026':'cantidad2026',
        'DESCRIP':    'Descrip',
        'CUODE':      'Cuode',
        'GCE3 ':      'GCE3',
        'CUCI3A5':    'CUCI3a5',
        'CUCISEC':    'CUCIsec',
        'CUCICAP':    'CUCIcap',
        'CUCIGRUP':   'CUCIgrup',
        'CUCISUBG':   'CUCIsubg',
        'CUCIPP':     'CUCIpp',
    })

    # Reordenar columnas en el mismo orden que el SAS
    cols_sas = [
        'posara', 'PARTE', 'CAPITULO', 'deptor', 'pais', 'MES', '_FREQ_',
        'FOBDOL2025', 'KNETO2025', 'fobpes2025', 'cantidad2025',
        'FOBDOL2026', 'KNETO2026', 'fobpes2026', 'cantidad2026',
        'Descrip', 'Cuode', 'CIIU3', 'CPC', 'CUCI2', 'CUCI3', 'GCE3',
        'CIIU32', 'CPC2', 'CUCI3a5', 'CUCIsec', 'CUCIcap', 'CUCIgrup',
        'CUCIsubg', 'CUCIpp', 'CPCV2', 'CIIU4', 'CIIU4_4', 'CIIU3_4',
    ]
    eje = eje[[c for c in cols_sas if c in eje.columns]]

    # Exportar a Excel (equivale al PROC EXPORT)
    print(f"Exportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        eje.to_excel(writer, sheet_name='EJE', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
