"""
Migración a Python: Lectura de base de importaciones archivo .FIN
Equivalente al programa SAS 'Lectura Impo base .FIN.sas'.

Lee un archivo .FIN de importaciones (LRECL=378), genera totales mensuales
y muestra el resumen por MES (equivale a PROC SUMMARY + PROC PRINT).
Dependencias: pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

ARCHIVO_FIN    = Path(r'\\systema44\Migracion\E100\M10126.FIN')
ARCHIVO_SALIDA = Path(r'D:\COMEX\IMPO\Resultados\Lectura_Impo_FIN.xlsx')

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO (LRECL=378)
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS col_ini-col_fin o @pos
# =============================================================================

COLSPECS = [
    (0,  4),    # MES      1-4  (lectura por rango de columnas)
    (22, 26),   # REGIMEN  @23 $4.
    (71, 81),   # NABAN    @72 $10.  (POSARA, 10 dígitos)
    (29, 42),   # PBK      @30 13.2
    (42, 55),   # PNK      @43 13.2
    (89, 100),  # VAFOB    @90 11.2
    (100, 111), # FLETE    @101 11.2
    (111, 124), # VACID    @112 13.2
    (304, 317), # SEGUROS  @305 13.2
    (317, 330), # GASTOS   @318 13.2
]

NOMBRES = ['MES', 'REGIMEN', 'NABAN', 'PBK', 'PNK', 'VAFOB', 'FLETE', 'VACID', 'SEGUROS', 'GASTOS']

COLS_STR = {'REGIMEN', 'NABAN'}

# Campos con 2 decimales implícitos (informat SAS w.2)
COLS_DEC2 = ['PBK', 'PNK', 'VAFOB', 'FLETE', 'VACID', 'SEGUROS', 'GASTOS']

VARS_SUMA = ['PNK', 'PBK', 'VAFOB', 'VACID', 'FLETE', 'SEGUROS', 'GASTOS']

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


def leer_fin(archivo) -> pd.DataFrame:
    """Lee el archivo .FIN de ancho fijo (equivale a DATA IMPO + INFILE DATOS)."""
    df = pd.read_fwf(
        archivo,
        colspecs=COLSPECS,
        names=NOMBRES,
        dtype={c: str for c in COLS_STR},
        header=None,
        encoding='latin-1',
    )

    for col in NOMBRES:
        if col not in COLS_STR:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    for col in COLS_DEC2:
        df[col] = _aplicar_decimal_implicito(df[col])

    df['NABAN']   = df['NABAN'].str.strip().str.zfill(10)
    df['REGIMEN'] = df['REGIMEN'].str.strip()

    return df


def resumir_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Equivale a:
      PROC SUMMARY DATA=IMPO;
        CLASS MES;
        VAR PNK PBK VAFOB VACID FLETE SEGUROS GASTOS;
        OUTPUT OUT=TOTAL SUM=;
    Genera una fila por MES más una fila de TOTAL GENERAL (_FREQ_=0).
    """
    por_mes = (
        df.groupby('MES', as_index=False, dropna=False)[VARS_SUMA]
        .sum()
    )
    por_mes['_FREQ_'] = df.groupby('MES')['MES'].transform('count').groupby(df['MES']).first().values

    total_gen = df[VARS_SUMA].sum().to_frame().T
    total_gen['MES']    = None
    total_gen['_FREQ_'] = len(df)

    return pd.concat([por_mes, total_gen], ignore_index=True)


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("=== LECTURA BASE IMPORTACIONES .FIN ===\n")

    print(f"Leyendo archivo: {ARCHIVO_FIN}")
    impo = leer_fin(ARCHIVO_FIN)
    print(f"Registros leídos: {len(impo):,}\n")

    # Equivale a PROC SUMMARY + PROC PRINT
    total = resumir_por_mes(impo)

    print("Resumen mensual:")
    pd.set_option('display.float_format', '{:,.2f}'.format)
    print(total.to_string(index=False))

    # Guardar a Excel
    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        total.to_excel(writer, sheet_name='TOTAL_MES', index=False)
        impo.to_excel(writer, sheet_name='DATOS', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
