"""Microtaller 23 · Análisis de rendimiento académico.

Lee los datos abiertos OULAD (Open University, Kuzilek et al., 2017), los limpia,
arma la base por matrícula, calcula los indicadores y deja listos:
  datos/base_matriculas.csv      una fila por estudiante y asignatura
  datos/rendimiento.db           la misma base en SQLite (matricula, evaluacion, nota)
  tablero/datos.js               agregados que usa el tablero interactivo
  analisis/resultados.json       cifras para el documento y las conclusiones

Uso:  python3 analisis/analisis_rendimiento.py
"""
import json
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
D = RAIZ / "datos" / "oulad"

AREA = {"AAA": "Ciencias sociales", "BBB": "Ciencias sociales", "GGG": "Ciencias sociales",
        "CCC": "Ciencias e ingeniería", "DDD": "Ciencias e ingeniería",
        "EEE": "Ciencias e ingeniería", "FFF": "Ciencias e ingeniería"}
COHORTE = {"2013B": "2013 febrero", "2013J": "2013 octubre",
           "2014B": "2014 febrero", "2014J": "2014 octubre"}
RESULTADO = {"Distinction": "Distinción", "Pass": "Aprobó", "Fail": "Reprobó", "Withdrawn": "Se retiró"}
ORDEN_RES = ["Distinción", "Aprobó", "Reprobó", "Se retiró"]
FORMACION = {"No Formal quals": "Sin estudios formales",
             "Lower Than A Level": "Secundaria incompleta",
             "A Level or Equivalent": "Secundaria completa",
             "HE Qualification": "Estudios superiores",
             "Post Graduate Qualification": "Posgrado"}
ORDEN_FORM = list(FORMACION.values())
RANGOS_NOTA = ["No la entregó", "Menos de 40", "40 a 59", "60 a 79", "80 a 100"]
DIA_CORTE = 28  # primeras cuatro semanas


def rango_nota(x):
    if pd.isna(x):
        return RANGOS_NOTA[0]
    if x < 40:
        return RANGOS_NOTA[1]
    if x < 60:
        return RANGOS_NOTA[2]
    if x < 80:
        return RANGOS_NOTA[3]
    return RANGOS_NOTA[4]


# ---------------------------------------------------------------- 1. carga
info = pd.read_csv(D / "studentInfo.csv")
ev = pd.read_csv(D / "assessments.csv", na_values=["?", ""])
sa = pd.read_csv(D / "studentAssessment.csv", na_values=["?", ""])
reg = pd.read_csv(D / "studentRegistration.csv", na_values=["?", ""])
CLAVE = ["code_module", "code_presentation", "id_student"]

limpieza = {
    "matriculas_leidas": len(info),
    "matriculas_duplicadas": int(info.duplicated(CLAVE).sum()),
    "notas_leidas": len(sa),
    "notas_sin_puntaje": int(sa["score"].isna().sum()),
    "notas_fuera_de_rango": int(((sa["score"] < 0) | (sa["score"] > 100)).sum()),
    "notas_trasladadas_de_intento_anterior": int(sa["is_banked"].sum()),
    "examenes_sin_fecha": int(ev.loc[ev["assessment_type"] == "Exam", "date"].isna().sum()),
    "imd_vacio": int(info["imd_band"].isna().sum()),
}

# ---------------------------------------------------------------- 2. notas por matrícula
sa = sa.dropna(subset=["score"]).merge(ev, on="id_assessment", how="left")
notas = sa.groupby(CLAVE).agg(evaluaciones_entregadas=("score", "size"),
                              promedio_evaluaciones=("score", "mean")).reset_index()

# orden de las evaluaciones del semestre (sin el examen) dentro de cada asignatura y cohorte
cont = ev[ev["assessment_type"] != "Exam"].sort_values(["code_module", "code_presentation", "date", "id_assessment"])
cont["orden"] = cont.groupby(["code_module", "code_presentation"]).cumcount() + 1
primera = cont[cont["orden"] == 1][["code_module", "code_presentation", "id_assessment", "date"]]
primera = primera.rename(columns={"id_assessment": "id_primera", "date": "dia_primera"})
p1 = sa.merge(primera, left_on=["code_module", "code_presentation", "id_assessment"],
              right_on=["code_module", "code_presentation", "id_primera"])[CLAVE + ["score"]]
p1 = p1.rename(columns={"score": "nota_primera_evaluacion"})

# ---------------------------------------------------------------- 3. actividad en la plataforma (primeras 4 semanas)
clics = []
for trozo in pd.read_csv(D / "studentVle.csv", usecols=CLAVE + ["date", "sum_click"],
                         dtype={"code_module": "category", "code_presentation": "category",
                                "id_student": "int32", "date": "int16", "sum_click": "int32"},
                         chunksize=2_000_000):
    t = trozo[trozo["date"] < DIA_CORTE]
    clics.append(t.groupby(CLAVE, observed=True)["sum_click"].sum())
clics = pd.concat(clics).groupby(level=[0, 1, 2]).sum().rename("clics_primeras_4_semanas").reset_index()
clics["code_module"] = clics["code_module"].astype(str)
clics["code_presentation"] = clics["code_presentation"].astype(str)

# ---------------------------------------------------------------- 4. base por matrícula
b = info.merge(reg[CLAVE + ["date_registration", "date_unregistration"]], on=CLAVE, how="left")
b = b.merge(notas, on=CLAVE, how="left").merge(p1, on=CLAVE, how="left").merge(clics, on=CLAVE, how="left")
b = b.merge(primera[["code_module", "code_presentation", "dia_primera"]], on=["code_module", "code_presentation"])
b["evaluaciones_entregadas"] = b["evaluaciones_entregadas"].fillna(0).astype(int)
b["clics_primeras_4_semanas"] = b["clics_primeras_4_semanas"].fillna(0).astype(int)

base = pd.DataFrame({
    "asignatura": b["code_module"],
    "area": b["code_module"].map(AREA),
    "cohorte": b["code_presentation"].map(COHORTE),
    "id_estudiante": b["id_student"],
    "genero": b["gender"].map({"M": "Hombre", "F": "Mujer"}),
    "formacion_previa": b["highest_education"].map(FORMACION),
    "rango_edad": b["age_band"].map({"0-35": "Menos de 35", "35-55": "35 a 55", "55<=": "55 o más"}),
    "intentos_previos": b["num_of_prev_attempts"],
    "creditos": b["studied_credits"],
    "discapacidad": b["disability"].map({"Y": "Sí", "N": "No"}),
    "nivel_desempeno": b["final_result"].map(RESULTADO),
    "aprobo": b["final_result"].isin(["Pass", "Distinction"]).map({True: "Sí", False: "No"}),
    "evaluaciones_entregadas": b["evaluaciones_entregadas"],
    "promedio_evaluaciones": b["promedio_evaluaciones"].round(1),
    "nota_primera_evaluacion": b["nota_primera_evaluacion"],
    "clics_primeras_4_semanas": b["clics_primeras_4_semanas"],
    "dia_retiro": b["date_unregistration"],
    # sigue inscrito el día de la primera evaluación y al terminar la semana 4
    "activo_en_primera_evaluacion": ~(b["date_unregistration"] < b["dia_primera"]),
    "activo_semana_4": ~(b["date_unregistration"] < DIA_CORTE),
})
assert base["nivel_desempeno"].notna().all() and base["formacion_previa"].notna().all()
base.sort_values(["asignatura", "cohorte", "id_estudiante"]).to_csv(RAIZ / "datos" / "base_matriculas.csv", index=False)

# SQLite: matrícula, evaluación y nota
with sqlite3.connect(RAIZ / "datos" / "rendimiento.db") as con:
    base.to_sql("matricula", con, if_exists="replace", index=False)
    e = ev.rename(columns={"code_module": "asignatura", "code_presentation": "cohorte_codigo", "id_assessment": "id_evaluacion",
                           "assessment_type": "tipo", "date": "dia", "weight": "peso"})
    e["tipo"] = e["tipo"].map({"TMA": "Trabajo calificado por tutor", "CMA": "Prueba calificada por computador", "Exam": "Examen final"})
    e.to_sql("evaluacion", con, if_exists="replace", index=False)
    sa[["id_assessment", "id_student", "date_submitted", "is_banked", "score"]].rename(columns={
        "id_assessment": "id_evaluacion", "id_student": "id_estudiante", "date_submitted": "dia_entrega",
        "is_banked": "trasladada", "score": "nota"}).to_sql("nota", con, if_exists="replace", index=False)

# ---------------------------------------------------------------- 5. indicadores
def tasas(df):
    n = len(df)
    c = df["nivel_desempeno"].value_counts()
    return {"n": n, **{r: round(100 * c.get(r, 0) / n, 1) for r in ORDEN_RES},
            "aprobacion": round(100 * (df["aprobo"] == "Sí").mean(), 1)}

por_asig = {a: {**tasas(g), "promedio": round(g["promedio_evaluaciones"].mean(), 1), "area": AREA[a]}
            for a, g in base.groupby("asignatura")}
critica = min(por_asig, key=lambda a: por_asig[a]["aprobacion"])

ok1 = base[base["activo_en_primera_evaluacion"]].copy()
ok1["rango_primera"] = ok1["nota_primera_evaluacion"].map(rango_nota)
patron_primera = {r: {**tasas(g)} for r, g in ok1.groupby("rango_primera")}

ok4 = base[base["activo_semana_4"]].copy()
ok4["quintil_clics"], cortes = pd.qcut(ok4["clics_primeras_4_semanas"].rank(method="first"), 5, labels=False, retbins=True)
lim = ok4.groupby("quintil_clics")["clics_primeras_4_semanas"].agg(["min", "max"])
patron_clics = {int(q): {**tasas(g), "min": int(lim.loc[q, "min"]), "max": int(lim.loc[q, "max"])} for q, g in ok4.groupby("quintil_clics")}

form = {f: tasas(g) for f, g in base.groupby("formacion_previa")}
intentos = {("0" if k == 0 else "1" if k == 1 else "2 o más"): None for k in [0, 1, 2]}
base["intentos_grupo"] = base["intentos_previos"].map(lambda k: "0" if k == 0 else "1" if k == 1 else "2 o más")
intentos = {k: tasas(g) for k, g in base.groupby("intentos_grupo")}

evo = sa[sa["assessment_type"] != "Exam"].merge(cont[["id_assessment", "orden"]], on="id_assessment")
inscritos = base.groupby("asignatura").size()

res = {
    "fuente": "Open University Learning Analytics Dataset (Kuzilek et al., 2017), CC BY 4.0",
    "limpieza": limpieza,
    "general": {**tasas(base), "promedio": round(base["promedio_evaluaciones"].mean(), 1),
                "estudiantes_unicos": int(base["id_estudiante"].nunique()),
                "sin_ninguna_nota": int((base["evaluaciones_entregadas"] == 0).sum())},
    "por_asignatura": por_asig,
    "asignatura_critica": critica,
    "patron_primera_evaluacion": patron_primera,
    "patron_clics_4_semanas": patron_clics,
    "por_formacion": form,
    "por_intentos_previos": intentos,
    "promedio_por_nivel": {r: round(g["promedio_evaluaciones"].mean(), 1) for r, g in base.groupby("nivel_desempeno")},
}
(RAIZ / "analisis" / "resultados.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")

# ---------------------------------------------------------------- 6. agregados para el tablero
# clave = asignatura | cohorte | formación; el tablero suma las claves que dejen pasar los filtros
asigs = sorted(base["asignatura"].unique())
cohs = list(COHORTE.values())
ok1 = ok1.set_index(ok1.index)
grupos = []
base["rango_primera"] = base["nota_primera_evaluacion"].map(rango_nota)
base["q_clics"] = np.nan
base.loc[ok4.index, "q_clics"] = ok4["quintil_clics"]
evo = evo.merge(base[["asignatura", "id_estudiante", "cohorte", "formacion_previa"]].assign(
    code_module=base["asignatura"], code_presentation=base["cohorte"].map({v: k for k, v in COHORTE.items()})),
    left_on=["code_module", "code_presentation", "id_student"],
    right_on=["code_module", "code_presentation", "id_estudiante"])
max_orden = int(cont["orden"].max())
for (a, c, f), g in base.groupby(["asignatura", "cohorte", "formacion_previa"]):
    ge = evo[(evo["asignatura"] == a) & (evo["cohorte"] == c) & (evo["formacion_previa"] == f)]
    eo = ge.groupby("orden")["score"].agg(["size", "sum"])
    n_ev = int(cont[(cont["code_module"] == a) & (cont["code_presentation"].map(COHORTE) == c)]["orden"].max())
    g1 = g[g["activo_en_primera_evaluacion"]]
    g4 = g[g["activo_semana_4"]]
    grupos.append({
        "a": a, "c": c, "f": f, "n": len(g),
        "r": [int((g["nivel_desempeno"] == r).sum()) for r in ORDEN_RES],
        "sp": round(float(g["promedio_evaluaciones"].sum()), 1), "np": int(g["promedio_evaluaciones"].notna().sum()),
        # por orden de evaluación: [inscritos que la tenían, entregas, suma de notas]
        "ev": [[len(g) if k <= n_ev else 0, int(eo["size"].get(k, 0)), round(float(eo["sum"].get(k, 0)), 1)] for k in range(1, max_orden + 1)],
        "p1": [[int(((g1["rango_primera"] == rr) & (g1["aprobo"] == "Sí")).sum()), int((g1["rango_primera"] == rr).sum())] for rr in RANGOS_NOTA],
        "q4": [[int(((g4["q_clics"] == q) & (g4["aprobo"] == "Sí")).sum()), int((g4["q_clics"] == q).sum())] for q in range(5)],
    })
datos_tablero = {
    "asignaturas": [{"id": a, "area": AREA[a]} for a in asigs], "cohortes": cohs, "formaciones": ORDEN_FORM,
    "resultados": ORDEN_RES, "rangos_nota": RANGOS_NOTA,
    "quintiles": [f"{patron_clics[q]['min']} a {patron_clics[q]['max']}" for q in range(5)],
    "grupos": grupos, "max_orden": max_orden,
}
(RAIZ / "tablero" / "datos.js").write_text("window.DATOS = " + json.dumps(datos_tablero, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")

# ---------------------------------------------------------------- 7. resumen en consola
g = res["general"]
print(f"Matrículas: {g['n']} · estudiantes únicos: {g['estudiantes_unicos']} · sin ninguna nota: {g['sin_ninguna_nota']}")
print(f"Promedio general: {g['promedio']} · aprobación {g['aprobacion']} % · distinción {g['Distinción']} % · reprobó {g['Reprobó']} % · se retiró {g['Se retiró']} %")
print("Limpieza:", limpieza)
print("\nPor asignatura (aprobación, promedio, retiro, reprobó, n):")
for a in sorted(por_asig, key=lambda a: por_asig[a]["aprobacion"]):
    p = por_asig[a]
    print(f"  {a} {p['area'][:10]:10} aprob {p['aprobacion']:5} · prom {p['promedio']:5} · retiro {p['Se retiró']:5} · reprob {p['Reprobó']:5} · n {p['n']}")
print("Asignatura crítica:", critica)
print("\nAprobación según la primera evaluación (activos ese día):")
for r in RANGOS_NOTA:
    p = patron_primera.get(r)
    print(f"  {r:14} aprob {p['aprobacion']:5} % · n {p['n']}")
print("\nAprobación según clics en las primeras 4 semanas (activos al día 28):")
for q in range(5):
    p = patron_clics[q]
    print(f"  Q{q+1} {p['min']:5}-{p['max']:<6} aprob {p['aprobacion']:5} % · retiro {p['Se retiró']:5} % · n {p['n']}")
print("\nPor formación previa:", {k: v["aprobacion"] for k, v in form.items()})
print("Por intentos previos:", {k: (v["aprobacion"], v["n"]) for k, v in intentos.items()})
print("Promedio por nivel:", res["promedio_por_nivel"])
print(f"\ndatos.js: {len(grupos)} grupos · {(RAIZ / 'tablero' / 'datos.js').stat().st_size // 1024} KB")
