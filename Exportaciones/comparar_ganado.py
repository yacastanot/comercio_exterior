import pandas as pd

F_PY  = r"D:\COMEX\EXPO\Resultados\Exportaciones_GanadovsCarne2025-2026_NIT.xlsx"
F_SAS = r"D:\COMEX\EXPO\Resultados\Exportaciones_GanadovsCarne2025-2026_NIT_SAS.xlsx"

sas       = pd.read_excel(F_SAS, sheet_name="Sheet1")
py_eje    = pd.read_excel(F_PY, sheet_name="EJE")
py_det    = pd.read_excel(F_PY, sheet_name="TOTAL_DETALLE")
py_posara = pd.read_excel(F_PY, sheet_name="TOTAL_POSARA")

print(f"SAS Sheet1      : {len(sas):>4} filas x {len(sas.columns)} cols")
print(f"PY EJE          : {len(py_eje):>4} filas x {len(py_eje.columns)} cols")
print(f"PY TOTAL_DETALLE: {len(py_det):>4} filas x {len(py_det.columns)} cols")
print(f"PY TOTAL_POSARA : {len(py_posara):>4} filas x {len(py_posara.columns)} cols")

# Columnas
cols_sas = set(sas.columns)
cols_det  = set(py_det.columns)
print(f"\nEn SAS no en DETALLE Python : {cols_sas - cols_det}")
print(f"En DETALLE Python no en SAS : {cols_det - cols_sas}")

# Comparar filas por clave de agrupacion
key = ["POSARA", "MES", "PAIS", "MODALIDAD", "LUGSAL", "NIT", "RAZON"]

sas_k = sas[key].fillna("").astype(str).apply(lambda r: "|".join(r.str.strip()), axis=1)
det_k = py_det[key].fillna("").astype(str).apply(lambda r: "|".join(r.str.strip()), axis=1)

solo_sas = set(sas_k) - set(det_k)
solo_py  = set(det_k) - set(sas_k)
comunes  = set(sas_k) & set(det_k)

print(f"\n--- Filas por clave {key} ---")
print(f"  Solo en SAS    : {len(solo_sas)}")
print(f"  Solo en Python : {len(solo_py)}")
print(f"  Comunes        : {len(comunes)}")

# Ejemplos de filas solo en SAS
if solo_sas:
    print("\nEjemplos filas SOLO en SAS:")
    mask = sas_k.isin(solo_sas)
    print(sas.loc[mask, key + ["FOBDOL2026","FOBDOL2025"]].head(5).to_string(index=False))

# Ejemplos de filas solo en Python
if solo_py:
    print("\nEjemplos filas SOLO en Python:")
    mask = det_k.isin(solo_py)
    print(py_det.loc[mask, key + ["FOBDOL2026","FOBDOL2025"]].head(5).to_string(index=False))

# Comparar valores en filas comunes
print("\n--- Diferencias en valores (filas comunes) ---")
VARS = ["FOBDOL2026", "KNETO2026", "FOBPES2026", "CANTIDAD2026",
        "FOBDOL2025", "KNETO2025", "FOBPES2025", "CANTIDAD2025"]

sas2 = sas.copy()
det2 = py_det.copy()
sas2["_key"] = sas_k
det2["_key"] = det_k

merged = sas2[sas2["_key"].isin(comunes)].merge(
    det2[det2["_key"].isin(comunes)], on="_key", suffixes=("_SAS","_PY")
)

for v in VARS:
    col_s = f"{v}_SAS"
    col_p = f"{v}_PY"
    if col_s in merged.columns and col_p in merged.columns:
        diff = (pd.to_numeric(merged[col_s], errors="coerce") -
                pd.to_numeric(merged[col_p], errors="coerce")).abs()
        n_dif = (diff > 0.01).sum()
        print(f"  {v:15s}: {n_dif:>4} filas con diferencia > 0.01")
