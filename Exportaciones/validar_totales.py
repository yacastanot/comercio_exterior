"""
Valida que los totales de la base principal (CSV del AVA)
coincidan con los totales del programa Lectura Expo base DEF.
"""
import pandas as pd

F_BASE = r"D:\COMEX\EXPO\Resultados\Base_Exportaciones_2026.csv"
F_DEF  = r"D:\COMEX\EXPO\Resultados\Resumen_Expo_DEF_2026.xlsx"

VARS = ['KILO_BR3', 'KILO_NET', 'VAL_FOB', 'FOB_PES3']

# =============================================================================
# Leer base principal y calcular totales por mes
# =============================================================================
base = pd.read_csv(F_BASE, sep=';', encoding='utf-8-sig', keep_default_na=False)

for v in VARS:
    base[v] = pd.to_numeric(base[v], errors='coerce')

tot_base = base.groupby('MES_PROC')[VARS].sum().reset_index()
tot_base_total = pd.DataFrame(
    [['TOTAL'] + [base[v].sum() for v in VARS]],
    columns=['MES_PROC'] + VARS
)
tot_base = pd.concat([tot_base, tot_base_total], ignore_index=True)
tot_base['MES_PROC'] = tot_base['MES_PROC'].astype(str)

# =============================================================================
# Leer totales DEF
# =============================================================================
tot_def = pd.read_excel(F_DEF, sheet_name='TOTAL_MES')
tot_def['MES_PROC'] = tot_def['MES_PROC'].astype(str)

# =============================================================================
# Comparar
# =============================================================================
print(f"{'':=<80}")
print("TOTALES BASE PRINCIPAL (AVA filtrado)")
print(f"{'':=<80}")
print(tot_base.to_string(index=False, float_format=lambda x: f'{x:,.2f}'))

print(f"\n{'':=<80}")
print("TOTALES DEF")
print(f"{'':=<80}")
print(tot_def.to_string(index=False, float_format=lambda x: f'{x:,.2f}'))

print(f"\n{'':=<80}")
print("DIFERENCIAS (BASE - DEF)")
print(f"{'':=<80}")

merged = tot_base.merge(tot_def, on='MES_PROC', suffixes=('_BASE', '_DEF'))

hay_diferencias = False
for v in VARS:
    merged[f'DIFF_{v}'] = pd.to_numeric(merged[f'{v}_BASE'], errors='coerce') \
                        - pd.to_numeric(merged[f'{v}_DEF'],  errors='coerce')

for _, row in merged.iterrows():
    print(f"\n  Mes: {row['MES_PROC']}")
    for v in VARS:
        diff = row[f'DIFF_{v}']
        estado = 'OK' if abs(diff) < 0.01 else f'*** DIFERENCIA: {diff:+,.2f}'
        print(f"    {v:12s}: BASE={float(row[f'{v}_BASE']):>22,.2f}  "
              f"DEF={float(row[f'{v}_DEF']):>22,.2f}  {estado}")
        if abs(diff) >= 0.01:
            hay_diferencias = True

print(f"\n{'':=<80}")
if hay_diferencias:
    print("RESULTADO: Existen diferencias entre las dos fuentes.")
else:
    print("RESULTADO: Los totales coinciden exactamente.")
print(f"{'':=<80}")
