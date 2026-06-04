"""
Migración a Python: Lectura Expo base .DEF
Equivalente al programa SAS original del mismo nombre.

Lee archivos .DEF de exportaciones, agrupa por mes y genera
totales de peso bruto, peso neto, FOB dólares y FOB pesos.

Dependencias:
    pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

RUTA_DEF      = Path(r'\\systema44\Migracion\E110')
RUTA_RESULTADO = Path(r'D:\COMEX\EXPO\Resultados')
ARCHIVO_SALIDA = RUTA_RESULTADO / 'Resumen_Expo_DEF_2026.xlsx'

ARCHIVOS_DEF = [
    RUTA_DEF / 'M10126.DEF',
    RUTA_DEF / 'M10226.DEF',
    # Agregar meses adicionales aquí cuando estén disponibles
]

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO (LRECL=378)
# Posiciones: (inicio_0idx, fin_0idx)  ←  SAS @pos = Python pos-1
# Solo se leen los campos usados en el programa SAS
# =============================================================================

COLSPECS = [
    (0,   4),    # ANO_PROC  @1   4.
    (4,   6),    # MES_PROC  @5   2.
    (139, 149),  # POS_ARAN  @140 10.
    (179, 194),  # KILO_BR3  @180 15.2
    (194, 209),  # KILO_NET  @195 15.2
    (209, 224),  # VAL_FOB   @210 15.2
    (224, 244),  # FOB_PES3  @225 20.2
]

NOMBRES = ['ANO_PROC', 'MES_PROC', 'POS_ARAN', 'KILO_BR3', 'KILO_NET', 'VAL_FOB', 'FOB_PES3']

# Columnas con 2 decimales implícitos (SAS w.2):
# si el valor en el archivo no contiene punto decimal, se divide entre 100
COLS_IMPLIED_DEC2 = ['KILO_BR3', 'KILO_NET', 'VAL_FOB', 'FOB_PES3']

VARS_SUMA = ['KILO_BR3', 'KILO_NET', 'VAL_FOB', 'FOB_PES3']

# =============================================================================
# FUNCIONES
# =============================================================================

def _aplicar_decimal_implicito(serie: pd.Series) -> pd.Series:
    """
    Replica el informat SAS w.2 operando sobre el string crudo del archivo:
    si el valor ya tiene punto decimal se usa tal cual; si no, se divide entre 100.
    Debe recibir strings crudos (dtype=str en read_fwf) para no perder
    dígitos decimales en números de gran magnitud como FOB_PES3.
    """
    def _convertir(val):
        if pd.isna(val):
            return float('nan')
        s = str(val).strip()
        if not s:
            return float('nan')
        return float(s) if '.' in s else float(s) / 100

    return pd.to_numeric(serie.apply(_convertir), errors='coerce')


def leer_def(archivos: list) -> pd.DataFrame:
    """Lee y concatena archivos .DEF de ancho fijo (equivale a DATA EXPO + INFILE)."""
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
    for col in ['ANO_PROC', 'MES_PROC']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # POS_ARAN: conservar como string con ceros a la izquierda
    df['POS_ARAN'] = df['POS_ARAN'].str.strip().str.zfill(10)

    return df


def resumir_por_mes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Equivale al PROC SUMMARY con CLASS MES_PROC:
    genera una fila por mes más una fila de total general.
    """
    por_mes = df.groupby('MES_PROC', as_index=False)[VARS_SUMA].sum()

    fila_total = pd.DataFrame(
        [['TOTAL'] + [df[v].sum() for v in VARS_SUMA]],
        columns=['MES_PROC'] + VARS_SUMA,
    )

    return pd.concat([por_mes, fila_total], ignore_index=True)


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("Leyendo archivos .DEF...")
    df = leer_def(ARCHIVOS_DEF)
    print(f"  Total registros leídos: {len(df):,}")

    # Resumen por mes (equivale a PROC SUMMARY + PROC PRINT)
    resumen = resumir_por_mes(df)

    # Imprimir en consola (equivale a PROC PRINT FORMAT _NUMERIC_ 18.2)
    print("\n--- Resumen por mes (FORMAT 18.2) ---")
    print(resumen.to_string(
        index=False,
        float_format=lambda x: f'{x:,.2f}',
    ))

    # Exportar a Excel
    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        resumen.to_excel(writer, sheet_name='TOTAL_MES', index=False)
        df.to_excel(writer, sheet_name='EXPO_BASE', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
