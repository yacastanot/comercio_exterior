"""
Migración a Python del programa SAS: UNIÓN_EXPO 2026.sas
Lectura de archivos de exportaciones de ancho fijo, corrección de
códigos de país y generación de resumen FOB/Kilos netos por país.
"""

import pandas as pd
from pathlib import Path

# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------
RUTA_FUENTE = Path(r"\\systema75\migracion\UNION COMEX\UNION_COMEX_2026\UNION_EXPO")
RUTA_SALIDA = Path(r"D:\UNION\BASE_EXPO")

# Agregar o comentar los meses disponibles
ARCHIVOS = [
    RUTA_FUENTE / "M1ZF0126.txt",
    # RUTA_FUENTE / "M1ZF0226.txt",
    # RUTA_FUENTE / "M1ZF0326.txt",
    # RUTA_FUENTE / "M1ZF0426.txt",
    # RUTA_FUENTE / "M1ZF0526.txt",
    # RUTA_FUENTE / "M1ZF0626.txt",
    # RUTA_FUENTE / "M1ZF0726.txt",
    # RUTA_FUENTE / "M1ZF0826.txt",
    # RUTA_FUENTE / "M1ZF0926.txt",
    # RUTA_FUENTE / "M1ZF1026.txt",
    # RUTA_FUENTE / "M1ZF1126.txt",
    # RUTA_FUENTE / "M1ZF1226.txt",
]

# ---------------------------------------------------------------------------
# Layout del archivo de ancho fijo (equivalente al INPUT de SAS)
# colspecs: (inicio, fin) en base 0, fin excluido — igual que SAS @col pero -1
# ---------------------------------------------------------------------------
COLSPECS = [
    (0,   4),    # ANO_PROC
    (4,   6),    # MES_PROC
    (6,   10),   # NRO_ART
    (10,  12),   # COD_INCO
    (12,  14),   # COD_ADUA
    (14,  15),   # TIP_IDENT
    (15,  27),   # NIT
    (27,  29),   # TIP_USUA
    (29,  34),   # COD_USUA
    (34,  36),   # TIP_EXPO
    (36,  41),   # DPTOMPIO
    (41,  44),   # COD_PAIS
    (44,  47),   # COD_PAI1
    (47,  67),   # CIU_PAIS
    (67,  83),   # AUT_EMB
    (83,  84),   # TIP_DECL
    (84,  86),   # COD_SAL1
    (86,  89),   # COD_SALI
    (89,  91),   # DPTO_PRO
    (91,  105),  # DEC_EXP
    (105, 109),  # ANO_EXP
    (109, 111),  # MES_EXP
    (111, 125),  # DEC_EXPO
    (125, 129),  # ANO_EXPO
    (129, 131),  # MES_EXPO
    (131, 135),  # COD_MIMP
    (135, 138),  # COD_MONE
    (138, 139),  # COD_VIA
    (139, 142),  # BANDERA
    (142, 143),  # COD_REGI
    (143, 146),  # COD_MODA
    (146, 147),  # FOR_PAGO
    (147, 148),  # COD_EMB
    (148, 149),  # COD_DAT
    (149, 150),  # CER_ORIG
    (150, 151),  # SIS_ESPE
    (151, 152),  # EXP_TRAN
    (152, 162),  # POS_ARAN
    (162, 166),  # N_SUPLEM
    (166, 170),  # N_COMPLE
    (170, 172),  # DPTO_ORI
    (172, 174),  # COD_MEDI
    (174, 177),  # CODUNI2
    (177, 192),  # CANT_UNI
    (192, 207),  # KILO_BR3
    (207, 222),  # KILO_NET
    (222, 237),  # VAL_FOB
    (237, 257),  # FOB_PES3
    (257, 272),  # VAL_ANAL
    (272, 287),  # FLETES3
    (287, 302),  # SEGURO3
    (302, 317),  # OTROSG3
    (317, 319),  # COD_AEMB
    (319, 323),  # ANO_EMB
    (323, 325),  # MES_EMB
    (325, 338),  # NUM_DECL
    (338, 342),  # ANO_DEC
    (342, 344),  # MES_DEC
    (344, 346),  # DIA_DEC
    (346, 406),  # RACIAL
    (406, 466),  # DIR_EXPO
    (466, 478),  # IDENTIFT
    (478, 538),  # NOM_APEL
    (538, 598),  # RACIALIM
    (598, 678),  # DIR_PDES
    (678, 681),  # ACUERDO
    (681, 684),  # CODZONA
    (684, 685),  # OPERA
    (685, 686),  # TIP_OPER
    (686, 687),  # TIP_USU
    (687, 690),  # COD_OPER
    (690, 700),  # NUM_FOR
]

NOMBRES = [
    "ANO_PROC", "MES_PROC", "NRO_ART",  "COD_INCO", "COD_ADUA", "TIP_IDENT",
    "NIT",      "TIP_USUA", "COD_USUA", "TIP_EXPO", "DPTOMPIO", "COD_PAIS",
    "COD_PAI1", "CIU_PAIS", "AUT_EMB",  "TIP_DECL", "COD_SAL1", "COD_SALI",
    "DPTO_PRO", "DEC_EXP",  "ANO_EXP",  "MES_EXP",  "DEC_EXPO", "ANO_EXPO",
    "MES_EXPO", "COD_MIMP", "COD_MONE", "COD_VIA",  "BANDERA",  "COD_REGI",
    "COD_MODA", "FOR_PAGO", "COD_EMB",  "COD_DAT",  "CER_ORIG", "SIS_ESPE",
    "EXP_TRAN", "POS_ARAN", "N_SUPLEM", "N_COMPLE", "DPTO_ORI", "COD_MEDI",
    "CODUNI2",  "CANT_UNI", "KILO_BR3", "KILO_NET", "VAL_FOB",  "FOB_PES3",
    "VAL_ANAL", "FLETES3",  "SEGURO3",  "OTROSG3",  "COD_AEMB", "ANO_EMB",
    "MES_EMB",  "NUM_DECL", "ANO_DEC",  "MES_DEC",  "DIA_DEC",  "RACIAL",
    "DIR_EXPO", "IDENTIFT", "NOM_APEL", "RACIALIM", "DIR_PDES", "ACUERDO",
    "CODZONA",  "OPERA",    "TIP_OPER", "TIP_USU",  "COD_OPER", "NUM_FOR",
]

# Columnas que deben leerse como texto ($ en SAS)
COLUMNAS_TEXTO = {
    "NIT", "COD_PAIS", "COD_PAI1", "CIU_PAIS", "AUT_EMB", "COD_SALI",
    "DEC_EXP", "DEC_EXPO", "COD_MIMP", "COD_MONE", "COD_MODA", "COD_EMB",
    "COD_DAT", "EXP_TRAN", "N_SUPLEM", "N_COMPLE", "CODUNI2", "NUM_DECL",
    "RACIAL", "DIR_EXPO", "IDENTIFT", "NOM_APEL", "RACIALIM", "DIR_PDES",
    "ACUERDO", "COD_OPER",
}

# Códigos de país que necesitan relleno con cero a la izquierda (SAS DATA ARREGLO)
CODIGOS_PAIS_RELLENAR = {
    "17", "23", "24", "27", "28", "32", "36", "40", "43", "44",
    "48", "52", "53", "56", "59", "60", "63", "68", "69", "72",
    "74", "76", "77", "80", "81", "83", "84", "87", "88", "90",
    "91", "92", "93", "97",
}


# ---------------------------------------------------------------------------
# Lectura de archivos (equivalente a DATA UNION_EXPO2026)
# ---------------------------------------------------------------------------
def leer_archivos(archivos: list[Path]) -> pd.DataFrame:
    dtype = {col: str for col in COLUMNAS_TEXTO}
    partes = []
    for ruta in archivos:
        df = pd.read_fwf(
            ruta,
            colspecs=COLSPECS,
            names=NOMBRES,
            dtype=dtype,
            encoding="latin-1",
            header=None,
        )
        partes.append(df)
        print(f"  Leído: {ruta.name}  ({len(df):,} registros)")
    return pd.concat(partes, ignore_index=True)


# ---------------------------------------------------------------------------
# Corrección de códigos de país (equivalente a DATA ARREGLO)
# ---------------------------------------------------------------------------
def corregir_codigos_pais(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["COD_PAIS1"] = df["COD_PAIS"].str.strip()
    mask = df["COD_PAIS1"].isin(CODIGOS_PAIS_RELLENAR)
    df.loc[mask, "COD_PAIS1"] = df.loc[mask, "COD_PAIS1"].str.zfill(3)
    df["AA"] = df["COD_PAIS1"].str.len()
    return df


# ---------------------------------------------------------------------------
# Resumen por país (equivalente a PROC SQL → CUADRO)
# ---------------------------------------------------------------------------
def generar_cuadro(df: pd.DataFrame) -> pd.DataFrame:
    cuadro = (
        df.groupby("COD_PAIS1", as_index=False)
        .agg(
            VAL_FOB=("VAL_FOB", "sum"),
            KILO_NET=("KILO_NET", "sum"),
        )
    )
    cuadro["VAL_FOB"]  = cuadro["VAL_FOB"]  / 1000
    cuadro["KILO_NET"] = cuadro["KILO_NET"] / 1000
    return cuadro


# ---------------------------------------------------------------------------
# Ejecución principal
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=== UNIÓN EXPO 2026 ===\n")

    print("Leyendo archivos fuente...")
    union_expo2026 = leer_archivos(ARCHIVOS)
    print(f"Total registros leídos: {len(union_expo2026):,}\n")

    print("Corrigiendo códigos de país...")
    arreglo = corregir_codigos_pais(union_expo2026)

    print("Generando cuadro resumen...")
    cuadro = generar_cuadro(arreglo)
    print(f"Países distintos: {len(cuadro)}\n")

    print("Cuadro resumen (VAL_FOB y KILO_NET en miles):")
    print(cuadro.to_string(index=False))

    # Guardar resultados
    salida_excel = RUTA_SALIDA / "UNION_EXPO_2026_CUADRO.xlsx"
    salida_base  = RUTA_SALIDA / "UNION_EXPO_2026_BASE.csv"

    print(f"\nGuardando resultados en {RUTA_SALIDA}...")
    cuadro.to_excel(salida_excel, index=False)
    arreglo.to_csv(salida_base, index=False, encoding="latin-1")

    print("Proceso finalizado correctamente.")
