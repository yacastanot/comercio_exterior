import pandas as pd

F_PY  = r"D:\COMEX\EXPO\Resultados\Base_Exportaciones_2026.csv"
F_SAS = r"D:\COMEX\EXPO\Resultados\COMER_Expo para generar Base CSV 2026.csv"

py  = pd.read_csv(F_PY,  dtype=str, encoding='utf-8-sig', sep=';', keep_default_na=False)
sas = pd.read_csv(F_SAS, dtype=str, encoding='latin-1',  sep=';', keep_default_na=False)

# Strip espacios
for c in py.columns:  py[c]  = py[c].str.strip()
for c in sas.columns: sas[c] = sas[c].str.strip()

print("=== COLUMNAS ===")
extra_py  = set(py.columns)  - set(sas.columns)
extra_sas = set(sas.columns) - set(py.columns)
print(f"Solo en Python : {extra_py  if extra_py  else 'ninguna'}")
print(f"Solo en SAS    : {extra_sas if extra_sas else 'ninguna'}")

cols = [c for c in sas.columns if c in py.columns]
py2, sas2 = py[cols].copy(), sas[cols].copy()

print(f"\n=== FILAS: Python={len(py2):,}  SAS={len(sas2):,} ===")

print("\n=== DIFERENCIAS POR COLUMNA ===")
diffs = {c: int((py2[c] != sas2[c]).sum()) for c in cols if (py2[c] != sas2[c]).any()}

if not diffs:
    print("Sin diferencias — archivos idénticos en columnas comunes.")
else:
    for col, n in sorted(diffs.items(), key=lambda x: -x[1]):
        print(f"  {col:20s}: {n:>8,} filas difieren")
    print(f"\nTotal columnas con diferencias: {len(diffs)}")

    print("\n=== EJEMPLOS (hasta 3 filas por columna) ===")
    for col in list(diffs.keys())[:10]:
        mask = py2[col] != sas2[col]
        ej = pd.DataFrame({"Python": py2.loc[mask, col],
                           "SAS":    sas2.loc[mask, col]}).head(3)
        print(f"\n--- {col} ---")
        print(ej.to_string())
