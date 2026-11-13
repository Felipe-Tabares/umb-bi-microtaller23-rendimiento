"""Genera un Excel FICTICIO, en escala de 0 a 5, para mostrar la carga de archivos del tablero.
No son datos reales: sirven solo para probar que el tablero lee una hoja de notas con otro formato."""
from pathlib import Path
import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
rng = np.random.default_rng(23)
filas = []
k = 0
for asig, dificultad in [("Cálculo Integral", 0.55), ("Programación", 0.25), ("Estadística", 0.40)]:
    for jornada in ("Diurna", "Nocturna"):
        for _ in range(30):
            k += 1
            compromiso = rng.uniform(0, 1) - (0.15 if jornada == "Nocturna" else 0)
            asist = int(np.clip(round(2 + 6 * compromiso + rng.normal(0, 1)), 0, 8))
            base = 2.2 + 2.6 * compromiso - dificultad + rng.normal(0, 0.35)
            notas = [float(np.clip(round(base + rng.normal(0, 0.45), 1), 0, 5)) for _ in range(3)]
            if compromiso < 0.12 and rng.uniform() < 0.8:          # se va después del primer corte
                notas = [notas[0], None, None]
            if compromiso < 0.05 and rng.uniform() < 0.5:          # ni siquiera entrega el primero
                notas = [None, None, None]
            definitiva = round(0.3 * notas[0] + 0.3 * notas[1] + 0.4 * notas[2], 1) if None not in notas else None
            filas.append({"estudiante": f"EST-{k:03d}", "asignatura": asig, "periodo": "2026-2", "jornada": jornada,
                          "corte_1": notas[0], "corte_2": notas[1], "corte_3": notas[2], "definitiva": definitiva,
                          "actividad_primer_mes": asist})
df = pd.DataFrame(filas)
salida = RAIZ / "entrega" / "Ejemplo-FICTICIO-notas-0-a-5.xlsx"
with pd.ExcelWriter(salida) as w:
    df.to_excel(w, sheet_name="Datos", index=False)
    pd.DataFrame({"Aviso": ["DATOS FICTICIOS generados con analisis/ejemplo_ficticio.py (semilla 23).",
                            "Solo sirven para probar la carga de archivos del tablero. No son estudiantes reales."]}
                 ).to_excel(w, sheet_name="Aviso", index=False)
# lo que el tablero debería mostrar, calculado aparte con pandas
ev = df[["corte_1", "corte_2", "corte_3"]]
p = ev.mean(axis=1)
nota = df["definitiva"].fillna(p)
r = np.where(nota.isna(), 3, np.where(nota >= 4.5, 0, np.where(nota >= 3, 1, 2)))
print(f"{len(df)} filas · promedio {p.mean():.1f} · aprobación {100*np.isin(r,[0,1]).mean():.1f} % · reprobó {100*(r==2).mean():.1f} % · retiro {100*(r==3).mean():.1f} %")
por = pd.DataFrame({"a": df.asignatura, "ok": np.isin(r, [0, 1])}).groupby("a")["ok"].mean().mul(100).round(1)
print("aprobación por asignatura:", por.to_dict())
