"""
Migración a Python: Expo para generar Base CSV 2026
Equivalente al programa SAS original del mismo nombre.

Lee los archivos .ava de exportaciones 2026, aplica los filtros de exclusión
y genera una base en formato CSV.

Dependencias:
    pip install pandas
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

RUTA_AVA       = Path(r'\\systema44\Migracion\E110')
RUTA_RESULTADO = Path(r'D:\COMEX\EXPO\Resultados')
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Base_Exportaciones_2026.csv'

ARCHIVOS_2026 = [
    RUTA_AVA / 'M10126.AVA',
    RUTA_AVA / 'M10226.AVA',
    # RUTA_AVA / 'M10326.AVA',
    # RUTA_AVA / 'M10426.AVA',
    # RUTA_AVA / 'M10526.AVA',
    # RUTA_AVA / 'M10626.AVA',
    # RUTA_AVA / 'M10726.AVA',
    # RUTA_AVA / 'M10826.AVA',
    # RUTA_AVA / 'M10926.AVA',
    # RUTA_AVA / 'M11026.AVA',
    # RUTA_AVA / 'M11126.AVA',
    # RUTA_AVA / 'M11226.AVA',
]

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO (LRECL=678)
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos = Python pos-1
# =============================================================================

COLSPECS = [
    (0,   4),    # ANO_PROC  @1   4.
    (4,   6),    # MES_PROC  @5   2.
    (6,   10),   # NRO_ART   @7   4.
    (10,  12),   # COD_INCO  @11  2.
    (12,  14),   # COD_ADUA  @13  2.
    (14,  15),   # TIP_IDENT @15  1.
    (15,  27),   # NIT       @16  $12.
    (27,  29),   # TIP_USUA  @28  2.
    (29,  34),   # COD_USUA  @30  5.
    (34,  36),   # TIP_EXPO  @35  2.
    (36,  41),   # DPTOMPIO  @37  5.
    (41,  44),   # COD_PAIS  @42  $3.
    (44,  47),   # COD_PAI1  @45  $3.
    (47,  67),   # CIU_PAIS  @48  $20.
    (67,  83),   # AUT_EMB   @68  $16.
    (83,  84),   # TIP_DECL  @84  1.
    (84,  86),   # COD_SAL1  @85  2.
    (86,  89),   # COD_SALI  @87  $3.
    (89,  91),   # DPTO_PRO  @90  2.
    (91,  105),  # DEC_EXP   @92  $14.
    (105, 109),  # ANO_EXP   @106 4.
    (109, 111),  # MES_EXP   @110 2.
    (111, 125),  # DEC_IMPO  @112 $14.
    (125, 129),  # ANO_IMPO  @126 4.
    (129, 131),  # MES_IMPO  @130 2.
    (131, 135),  # COD_MIMP  @132 $4.
    (135, 138),  # COD_MONE  @136 $3.
    (138, 139),  # COD_VIA   @139 1.
    (139, 142),  # BANDERA   @140 3.
    (142, 143),  # COD_REGI  @143 1.
    (143, 146),  # COD_MODA  @144 $3.
    (146, 147),  # FOR_PAGO  @147 1.
    (147, 148),  # COD_EMB   @148 $1.
    (148, 149),  # COD_DAT   @149 $1.
    (149, 150),  # CER_ORIG  @150 1.
    (150, 151),  # SIS_ESPE  @151 1.
    (151, 152),  # EXP_TRAN  @152 $1.
    (152, 162),  # POS_ARAN  @153 10.
    (162, 166),  # N_SUPLEM  @163 $4.
    (166, 170),  # N_COMPLE  @167 $4.
    (170, 172),  # DPTO_ORI  @171 2.
    (172, 174),  # COD_MEDI  @173 2.
    (174, 177),  # CODUNI2   @175 $3.
    (177, 192),  # CANT_UNI  @178 15.2
    (192, 207),  # KILO_BR3  @193 15.2
    (207, 222),  # KILO_NET  @208 15.2
    (222, 237),  # VAL_FOB   @223 15.2
    (237, 257),  # FOB_PES3  @238 20.2
    (257, 272),  # VAL_ANAL  @258 15.2
    (272, 287),  # FLETES3   @273 15.2
    (287, 302),  # SEGURO3   @288 15.2
    (302, 317),  # OTROSG3   @303 15.2
    (317, 319),  # COD_AEMB  @318 2.
    (319, 323),  # ANO_EMB   @320 4.
    (323, 325),  # MES_EMB   @324 2.
    (325, 338),  # NUM_DECL  @326 $13.
    (338, 342),  # ANO_DEC   @339 4.
    (342, 344),  # MES_DEC   @343 2.
    (344, 346),  # DIA_DEC   @345 2.
    (346, 406),  # RACIAL    @347 $60.
    (406, 466),  # DIR_EXPO  @407 $60.
    (466, 478),  # IDENTIFT  @467 $12.
    (478, 538),  # NOM_APEL  @479 $60.
    (538, 598),  # RACIALIM  @539 $60.
    (598, 678),  # DIR_PDES  @599 $80.
]

NOMBRES = [
    'ANO_PROC', 'MES_PROC', 'NRO_ART', 'COD_INCO', 'COD_ADUA', 'TIP_IDENT',
    'NIT', 'TIP_USUA', 'COD_USUA', 'TIP_EXPO', 'DPTOMPIO', 'COD_PAIS',
    'COD_PAI1', 'CIU_PAIS', 'AUT_EMB', 'TIP_DECL', 'COD_SAL1', 'COD_SALI',
    'DPTO_PRO', 'DEC_EXP', 'ANO_EXP', 'MES_EXP', 'DEC_IMPO', 'ANO_IMPO',
    'MES_IMPO', 'COD_MIMP', 'COD_MONE', 'COD_VIA', 'BANDERA', 'COD_REGI',
    'COD_MODA', 'FOR_PAGO', 'COD_EMB', 'COD_DAT', 'CER_ORIG', 'SIS_ESPE',
    'EXP_TRAN', 'POS_ARAN', 'N_SUPLEM', 'N_COMPLE', 'DPTO_ORI', 'COD_MEDI',
    'CODUNI2', 'CANT_UNI', 'KILO_BR3', 'KILO_NET', 'VAL_FOB', 'FOB_PES3',
    'VAL_ANAL', 'FLETES3', 'SEGURO3', 'OTROSG3', 'COD_AEMB', 'ANO_EMB',
    'MES_EMB', 'NUM_DECL', 'ANO_DEC', 'MES_DEC', 'DIA_DEC', 'RACIAL',
    'DIR_EXPO', 'IDENTIFT', 'NOM_APEL', 'RACIALIM', 'DIR_PDES',
]

# Columnas que se leen como texto
COLS_STR = {
    'NIT', 'COD_PAIS', 'COD_PAI1', 'CIU_PAIS', 'AUT_EMB', 'COD_SALI',
    'DEC_EXP', 'DEC_IMPO', 'COD_MIMP', 'COD_MONE', 'COD_MODA', 'COD_EMB',
    'COD_DAT', 'EXP_TRAN', 'N_SUPLEM', 'N_COMPLE', 'CODUNI2', 'NUM_DECL',
    'RACIAL', 'DIR_EXPO', 'IDENTIFT', 'NOM_APEL', 'RACIALIM', 'DIR_PDES',
    'POS_ARAN',
}

# Columnas con 2 decimales implícitos (SAS w.2)
COLS_IMPLIED_DEC2 = [
    'CANT_UNI', 'KILO_BR3', 'KILO_NET', 'VAL_FOB',
    'FOB_PES3', 'VAL_ANAL', 'FLETES3', 'SEGURO3', 'OTROSG3',
]

# =============================================================================
# FILTROS DE EXCLUSIÓN
# Equivalen a los IF ... THEN DELETE del SAS
# =============================================================================

# Modalidades de exportación a excluir directamente
MODA_EXCLUIR = {
    '001', '003', '105', '301', '302', '303', '304', '305',
    '306', '307', '399', '308', '309', '701', '500', '202',
}

# Modalidades de tráfico de exportación a excluir cuando FOR_PAGO = 2
MODA_EXCLUIR_FORPAGO2 = {'401', '402', '403'}

# Posiciones arancelarias a excluir (oro, títulos, otros)
POS_ARAN_EXCLUIR = {'7108200000', '4907002000', '4907003000'}

# Aerolíneas a excluir: NITs y posiciones arancelarias de aeronaves
# (Aerocivil y Avianca registran aeronaves que no son exportaciones)
NITS_EXCLUIR_AERONAVES = {
    # Aerocivil
    '890912462', '0890912462', '00890912462', '000890912462',
    '8909124622', '08909124622', '008909124622',
    # Avianca
    '890100577', '0890100577', '00890100577', '000890100577',
    '8901005776', '08901005776', '008901005776',
}
MODA_AERONAVES   = {'401', '402'}
POS_ARAN_AERONAVES = {'8802400000', '8802309000'}

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    """
    Replica el informat SAS w.2 operando sobre el string crudo del archivo:
    si el valor ya tiene punto decimal, se usa tal cual;
    si no, se divide entre 100.
    IMPORTANTE: recibir strings crudos, no floats pre-parseados, para no
    perder dígitos decimales en números de gran magnitud.
    """
    def _convertir(val):
        if pd.isna(val):
            return float('nan')
        s = str(val).strip()
        if not s:
            return float('nan')
        return float(s) if '.' in s else float(s) / 100

    return pd.to_numeric(serie.apply(_convertir), errors='coerce')


def _fmt_float(x) -> str:
    """
    Formato SAS para columnas float: elimina el sufijo '.0' en valores enteros
    y conserva hasta 2 decimales significativos (igual que SAS BEST format).
    """
    if pd.isna(x):
        return ''
    if x == int(x):
        return str(int(x))
    return f'{x:.2f}'.rstrip('0').rstrip('.')


def _best12(value) -> str:
    """
    Replica el formato SAS BEST12.: muestra hasta 12 caracteres.
    Para enteros: sin punto decimal. Para decimales: redondea para
    ajustarse a 12 chars y elimina ceros finales.
    """
    if pd.isna(value):
        return ''
    v = float(value)
    if v == int(v):
        return str(int(v))
    int_len = len(str(int(abs(v))))
    avail_dec = 11 - int_len          # 12 chars - 1 punto decimal - dígitos enteros
    if avail_dec <= 0:
        return str(int(round(v)))
    return f'{v:.{avail_dec}f}'.rstrip('0').rstrip('.')


def leer_ava(archivos: list) -> pd.DataFrame:
    """Lee y concatena archivos .ava de ancho fijo (equivale a DATA COMER + INFILE)."""
    partes = []
    for ruta in archivos:
        # Leer TODAS las columnas como string para que _aplicar_decimal_implicito
        # trabaje sobre el string crudo y no sobre un float ya parseado.
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
    # para que el string crudo llegue intacto a _aplicar_decimal_implicito.
    for col in COLS_IMPLIED_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    # Convertir el resto de columnas numéricas
    for col in NOMBRES:
        if col not in COLS_STR and col not in COLS_IMPLIED_DEC2:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Limpiar espacios en columnas string
    for col in COLS_STR:
        df[col] = df[col].fillna('').str.strip()

    # parte: SAS lee @153 parte 4. — los mismos 4 bytes que inician POS_ARAN,
    # interpretados como número (sin ceros iniciales).
    # Se calcula sobre el string crudo (con ceros), ANTES del lstrip de POS_ARAN.
    raw_pos = df['POS_ARAN'].str.strip()
    df['parte'] = pd.to_numeric(
        raw_pos.str[:4], errors='coerce'
    ).astype('Int64')

    # POS_ARAN: SAS lo lee como 10. (numérico), lo que elimina ceros iniciales.
    # Replicar ese comportamiento para coincidir con el CSV de SAS.
    df['POS_ARAN'] = raw_pos.str.lstrip('0')
    df['POS_ARAN'] = df['POS_ARAN'].where(df['POS_ARAN'] != '', '0')

    # Normalizar COD_MODA y NIT (quitar espacios)
    df['COD_MODA'] = df['COD_MODA'].str.strip()
    df['NIT']      = df['NIT'].str.strip()

    return df


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica los filtros de exclusión del SAS (IF ... THEN DELETE).
    Retorna solo los registros que deben conservarse.
    """
    n_inicial = len(df)

    # IF COD_MODA IN(...) THEN DELETE
    mask1 = df['COD_MODA'].isin(MODA_EXCLUIR)

    # IF COD_MODA IN('401','402','403') AND FOR_PAGO = 2 THEN DELETE
    mask2 = df['COD_MODA'].isin(MODA_EXCLUIR_FORPAGO2) & (df['FOR_PAGO'] == 2)

    # IF POS_ARAN IN(...) THEN DELETE
    mask3 = df['POS_ARAN'].isin(POS_ARAN_EXCLUIR)

    # IF NIT IN(...) AND COD_MODA IN('401','402') AND POS_ARAN IN(...) THEN DELETE
    # (cubre Aerocivil y Avianca en el mismo filtro)
    mask4 = (
        df['NIT'].isin(NITS_EXCLUIR_AERONAVES)
        & df['COD_MODA'].isin(MODA_AERONAVES)
        & df['POS_ARAN'].isin(POS_ARAN_AERONAVES)
    )

    df_filtrado = df[~(mask1 | mask2 | mask3 | mask4)].copy()

    print(f"  Registros originales : {n_inicial:>10,}")
    print(f"  Excluidos filtro 1   : {mask1.sum():>10,}  (modalidades excluidas)")
    print(f"  Excluidos filtro 2   : {mask2.sum():>10,}  (modalidad 401/402/403 + FOR_PAGO=2)")
    print(f"  Excluidos filtro 3   : {mask3.sum():>10,}  (posiciones arancelarias especiales)")
    print(f"  Excluidos filtro 4   : {mask4.sum():>10,}  (aeronaves Aerocivil/Avianca)")
    print(f"  Registros conservados: {len(df_filtrado):>10,}")

    return df_filtrado


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("Leyendo archivos 2026...")
    df = leer_ava(ARCHIVOS_2026)
    print(f"  Total registros leídos: {len(df):,}")

    print("\nAplicando filtros de exclusión...")
    df_comer = aplicar_filtros(df)

    # FOB_PES3 usa formato SAS BEST12.: redondea solo cuando supera 12 chars.
    df_comer['FOB_PES3'] = df_comer['FOB_PES3'].apply(_best12)

    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    df_comer.to_csv(ARCHIVO_SALIDA, index=False, encoding='utf-8-sig',
                    sep=';', float_format=_fmt_float)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
