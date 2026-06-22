"""
Migración a Python: Lectura de base Colombia Productiva
Equivalente al programa SAS 'Lectura Impo base Colombia Productiva.sas'.

Lee el archivo especial de Colombia Productiva (_SNR.ASU), genera:
  1. Totales mensuales (sin país)
  2. Desglose por mes y país de origen
  3. Tabla combinada ordenada (total primero, luego desglose)

Dependencias: pip install pandas openpyxl
"""

import pandas as pd
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================

ARCHIVO_ASU    = Path(r'\\systema44\Migracion\AVANCE IMPORTACIONES\COLOMBIA PRODUCTIVA 2026\M10126_COLOMBIA_PRODUCTIVA_SNR.ASU')
ARCHIVO_SALIDA = Path(r'D:\COMEX\IMPO\Resultados\Colombia_Productiva_2026.xlsx')

# =============================================================================
# DEFINICIÓN DEL ARCHIVO DE ANCHO FIJO
# Mismo layout que los .asu estándar, con los campos necesarios
# =============================================================================

COLSPECS = [
    (0,  4),    # FECH     @1  $4.   (YYMM)
    (6,  9),    # PAOR     @7  $3.   (país origen)
    (22, 26),   # REGIMEN  $23-26
    (29, 42),   # PBK      @30 13.2
    (42, 55),   # PNK      @43 13.2
    (71, 81),   # NABAN    @72 $10.  (POSARA)
    (89, 100),  # VAFODO   @90 11.2
    (100, 111), # FLETE    @101 11.2
    (111, 124), # VACID    @112 13.2
    (304, 317), # SEGUROS  @305 13.2
    (317, 330), # OTROSG   @318 13.2
]

NOMBRES = ['FECH', 'PAOR', 'REGIMEN', 'PBK', 'PNK', 'NABAN',
           'VAFODO', 'FLETE', 'VACID', 'SEGUROS', 'OTROSG']

COLS_STR = {'FECH', 'PAOR', 'REGIMEN', 'NABAN'}

COLS_DEC2 = ['PBK', 'PNK', 'VAFODO', 'FLETE', 'VACID', 'SEGUROS', 'OTROSG']

VARS_SUMA = ['PNK', 'PBK', 'VAFODO', 'VACID', 'FLETE', 'SEGUROS', 'OTROSG']

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


def leer_asu(archivo) -> pd.DataFrame:
    """Lee el archivo .ASU de ancho fijo (equivale a DATA IMPO + INFILE)."""
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

    # MES desde FECH (posiciones 3-4)
    df['MES']   = pd.to_numeric(df['FECH'].str[2:4], errors='coerce')
    df['PAOR']  = df['PAOR'].str.strip()
    df['NABAN'] = df['NABAN'].str.strip().str.zfill(10)

    return df


def total_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Paso 1: PROC SUMMARY NWAY; CLASS MES; → totales por mes.
    Agrega TOTAL = suma de todos los valores numéricos.
    """
    grp = df.groupby('MES', as_index=False, dropna=False)[VARS_SUMA].sum()
    grp['PAOR']  = 'TOTAL'
    grp['TOTAL'] = grp[VARS_SUMA].sum(axis=1)
    return grp


def total_paises(df: pd.DataFrame) -> pd.DataFrame:
    """
    Paso 2: PROC SUMMARY NWAY; CLASS MES PAOR; → desglose por mes y país.
    """
    grp = df.groupby(['MES', 'PAOR'], as_index=False, dropna=False)[VARS_SUMA].sum()
    grp['TOTAL'] = grp[VARS_SUMA].sum(axis=1)
    return grp


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def main():
    print("=== LECTURA BASE COLOMBIA PRODUCTIVA ===\n")

    print(f"Leyendo archivo: {ARCHIVO_ASU}")
    impo = leer_asu(ARCHIVO_ASU)
    print(f"Registros leídos: {len(impo):,}\n")

    # Paso 1: totales mensuales (PAOR = "TOTAL")
    tot_mes = total_mensual(impo)
    print(f"Totales mensuales: {len(tot_mes)} filas")

    # Paso 2: desglose por mes y país
    tot_paises = total_paises(impo)
    print(f"Desglose por país: {len(tot_paises)} filas")

    # Paso 3: unir (DATA TOTAL_FINAL; SET TOTAL_MES TOTAL_PAISES)
    cols_comunes = ['MES', 'PAOR', 'TOTAL'] + VARS_SUMA
    total_final = pd.concat(
        [tot_mes[cols_comunes], tot_paises[cols_comunes]],
        ignore_index=True
    )

    # Paso 4: ordenar por MES ASC, PAOR DESC (total primero porque 'TOTAL' > códigos numéricos)
    total_final = total_final.sort_values(
        ['MES', 'PAOR'], ascending=[True, False]
    ).reset_index(drop=True)

    # Paso 5: mostrar resultados (PROC PRINT)
    cols_mostrar = ['MES', 'PAOR'] + VARS_SUMA
    pd.set_option('display.float_format', '{:,.2f}'.format)
    print("\nResultados (primeras 20 filas):")
    print(total_final[cols_mostrar].head(20).to_string(index=False))

    # Guardar a Excel
    print(f"\nExportando a: {ARCHIVO_SALIDA}")
    ARCHIVO_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
        total_final.to_excel(writer, sheet_name='TOTAL_FINAL', index=False)
        tot_mes.to_excel(writer, sheet_name='TOTAL_MES', index=False)
        tot_paises.to_excel(writer, sheet_name='TOTAL_PAISES', index=False)

    print("Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
