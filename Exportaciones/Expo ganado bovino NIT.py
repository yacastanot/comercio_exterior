"""
Migración a Python: Expo ganado bovino NIT
Equivalente al programa SAS original del mismo nombre.

Lee archivos .ava de exportaciones 2025-2026, filtra posiciones arancelarias
de ganado bovino y carne, agrega por múltiples variables y exporta a Excel.

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
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Exportaciones_GanadovsCarne2025-2026_NIT.xlsx'
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
    (0,   4),    # FECHPRO   @1   4.
    (4,   6),    # MES       @5   2.
    (12,  14),   # ADU       @13  2.
    (15,  27),   # NIT       @16  $12.
    (36,  41),   # MUNICIPIO @37  5.
    (41,  44),   # PAIS      @42  3.
    (47,  67),   # CIUEXP    @48  $20.
    (84,  86),   # LUGSALN   @85  2.
    (86,  89),   # LUGSAL    @87  $3.
    (89,  91),   # DEPTPR    @90  2.
    (91,  105),  # DEXANTE   @92  14.
    (138, 139),  # VIA       @139 1.
    (139, 142),  # BANDERA   @140 3.
    (142, 143),  # REGIMEN   @143 1.
    (143, 146),  # MODALIDAD @144 3.
    (146, 147),  # FORPA4    @147 1.
    (152, 162),  # POSARA    @153 10.  (string para preservar ceros)
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

COLS_STR = {
    'NIT', 'CIUEXP', 'LUGSAL', 'POSARA', 'DEX', 'FECHDEC',
    'FECHEMB', 'FECHPRO', 'RAZON', 'MODALIDAD',
}

# Columnas con 2 decimales implícitos (SAS w.2)
COLS_IMPLIED_DEC2 = ['CANTIDAD', 'KNETO', 'FOBDOL', 'FOBPES']

# Variables de agrupación del macro %CALC
# Equivale a: %CALC(posara mes pais modalidad lugsal NIT RAZON)
GRUPOS_DETALLE = ['POSARA', 'MES', 'PAIS', 'MODALIDAD', 'LUGSAL', 'NIT', 'RAZON']

VARS_SUMA = ['FOBDOL', 'KNETO', 'FOBPES', 'CANTIDAD']

# =============================================================================
# POSICIONES ARANCELARIAS DE GANADO BOVINO Y CARNE
# Los valores SAS son numéricos de 9 dígitos → se normalizan a 10 con zfill
# =============================================================================

POSARA_BOVINO = {
    '0102299020',  # Ganado bovino vivo
    '0201300010',  # Carne bovina fresca/refrigerada
    '0201300090',
    '0201300093',
    '0201300094',
    '0202200000',  # Carne bovina congelada
    '0202300010',
    '0202300090',
    '0202300093',
    '0202300094',
    '0202300095',
}

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    """
    Replica el informat SAS w.2 operando sobre el string crudo del archivo:
    si el valor ya tiene punto decimal se usa tal cual; si no, se divide entre 100.
    Debe recibir strings crudos (dtype=str en read_fwf) para no perder
    dígitos decimales en números de gran magnitud.
    """
    def _convertir(val):
        if pd.isna(val):
            return float('nan')
        s = str(val).strip()
        if not s:
            return float('nan')
        return float(s) if '.' in s else float(s) / 100

    return pd.to_numeric(serie.apply(_convertir), errors='coerce')


def leer_ava(archivos: list) -> pd.DataFrame:
    """Lee y concatena archivos .ava de ancho fijo."""
    partes = []
    for ruta in archivos:
        # Leer todo como string para que _aplicar_decimal_implicito trabaje
        # sobre el string crudo y no sobre un float ya parseado.
        df = pd.read_fwf(
            ruta,
            colspecs=COLSPECS,
            names=NOMBRES,
            dtype=str,
            header=None,
            encoding='latin-1',
        )
        partes.append(df)

    df = pd.concat(partes, ignore_index=True)

    # Decimales implícitos (SAS w.2): aplicar ANTES de pd.to_numeric
    for col in COLS_IMPLIED_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    # Convertir columnas numéricas restantes
    for col in NOMBRES:
        if col not in COLS_STR and col not in COLS_IMPLIED_DEC2:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Normalizar POSARA a 10 dígitos con ceros a la izquierda
    df['POSARA'] = df['POSARA'].str.strip().str.zfill(10)

    # Derivar campos arancelarios desde POSARA
    df['CAPITULO'] = df['POSARA'].str[:2].astype(int)
    df['PARTE']    = df['POSARA'].str[:6]   # substr(posara,1,6)
    df['SA']       = df['POSARA'].str[:6]

    # Normalizar campos de texto usados como agrupadores
    df['NIT']       = df['NIT'].str.strip()
    df['RAZON']     = df['RAZON'].str.strip()
    df['LUGSAL']    = df['LUGSAL'].str.strip()
    df['MODALIDAD'] = df['MODALIDAD'].str.strip()

    # Indicador TRA: excluye productos minero-energéticos
    posara_tra = {'0901110000', '0901111000', '0901119000', '7202600000'}
    parte_int  = df['POSARA'].str[:4].astype(int)
    mask_tra   = (
        df['POSARA'].isin(posara_tra)
        | parte_int.between(2701, 2704)
        | parte_int.between(2709, 2715)
    )
    df['TRA'] = 2
    df.loc[mask_tra, 'TRA'] = 1

    return df


def calc(df2025: pd.DataFrame, df2026: pd.DataFrame, grupo: list) -> pd.DataFrame:
    """
    Equivale al macro SAS %CALC:
    agrega cada año por 'grupo', hace merge externo y filtra ganado/carne.
    El merge es TOTAL2026 + TOTAL2025 (orden del SAS original).
    """
    tot25 = (
        df2025.groupby(grupo, as_index=False)[VARS_SUMA]
        .sum()
        .rename(columns={v: f'{v}2025' for v in VARS_SUMA})
    )
    tot26 = (
        df2026.groupby(grupo, as_index=False)[VARS_SUMA]
        .sum()
        .rename(columns={v: f'{v}2026' for v in VARS_SUMA})
    )
    # En el SAS: merge TOTAL2026 TOTAL2025 (2026 a la izquierda)
    merged = pd.merge(tot26, tot25, on=grupo, how='outer')

    # Filtrar solo ganado bovino y carne
    return merged[merged['POSARA'].isin(POSARA_BOVINO)].copy()


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

    # Agregación detallada: por posara + mes + pais + modalidad + lugsal + NIT + RAZON
    # Equivale a: %CALC(posara mes pais modalidad lugsal NIT RAZON)
    print("  Calculando TOTAL detalle (posara × mes × pais × modalidad × lugsal × NIT × RAZON)...")
    total_detalle = calc(df2025, df2026, GRUPOS_DETALLE)

    # Agregación solo por posara (para el cruce con corre62 en el paso EJE)
    print("  Calculando TOTAL por posara...")
    total_posara = calc(df2025, df2026, ['POSARA'])

    # Leer tabla de referencia corre62
    print("Leyendo tabla de referencia corre62...")
    corre62 = pd.read_excel(CORRE62_XLSX)
    corre62.columns = corre62.columns.str.upper()
    corre62['POSARA'] = corre62['POSARA'].astype(str).str.strip().str.zfill(10)

    # EJE: TOTAL_DETALLE + columnas de corre62 (merge por POSARA)
    # Equivale a: DATA expo.EJE; MERGE expo.totalposara(IN=A) EXPO.corre62; BY posara; IF A;
    # En SAS, "totalposara" es el resultado del %CALC con todos los grupos (= total_detalle).
    eje = pd.merge(total_detalle, corre62, on='POSARA', how='left')

    # Exportar a Excel
    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        eje.to_excel(writer, sheet_name='EJE', index=False)
        total_detalle.to_excel(writer, sheet_name='TOTAL_DETALLE', index=False)
        total_posara.to_excel(writer, sheet_name='TOTAL_POSARA', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
