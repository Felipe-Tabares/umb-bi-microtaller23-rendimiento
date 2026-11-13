"""Arma el libro de Excel del microtaller 23 (rendimiento académico, OULAD).

Uso:  /usr/local/bin/python3 -I analisis/armar_excel.py
Sale: entrega/Microtaller23-Rendimiento-Academico.xlsx (sin valores en caché;
      se abre y se guarda con Microsoft Excel para que queden calculados).
"""
import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = RAIZ / "entrega" / "Microtaller23-Rendimiento-Academico.xlsx"
RES = json.loads((RAIZ / "analisis" / "resultados.json").read_text(encoding="utf-8"))

FUENTE = "Calibri"
NEGRO = "000000"
F_NORMAL = Font(name=FUENTE, size=11, color=NEGRO)
F_NEGRITA = Font(name=FUENTE, size=11, color=NEGRO, bold=True)
F_TITULO = Font(name=FUENTE, size=14, color=NEGRO, bold=True)
F_NOTA = Font(name=FUENTE, size=10, color=NEGRO, italic=True)
GRIS = PatternFill("solid", fgColor="D9D9D9")
CLARO = PatternFill("solid", fgColor="F2F2F2")
ENTRADA = PatternFill("solid", fgColor="FFF2CC")
LINEA = Side(style="thin", color="808080")
BORDE = Border(left=LINEA, right=LINEA, top=LINEA, bottom=LINEA)
PCT = "0.0 %"
UN_DEC = "0.0"
MILES = "#,##0"
AJUSTE = Alignment(wrap_text=True, vertical="top")

ASIGNATURAS = ["AAA", "BBB", "CCC", "DDD", "EEE", "FFF", "GGG"]
NIVELES = ["Distinción", "Aprobó", "Reprobó", "Se retiró"]
RANGOS = ["No la entregó", "Menos de 40", "40 a 59", "60 a 79", "80 a 100"]

# ------------------------------------------------------------------ datos
base = pd.read_csv(RAIZ / "datos" / "base_matriculas.csv")
N = len(base)
ULT = N + 1  # última fila con datos en la hoja Base
COL = {c: get_column_letter(i + 1) for i, c in enumerate(base.columns)}


def rango(col):
    """Rango absoluto de una columna de la hoja Base, por ejemplo Base!$A$2:$A$32594."""
    letra = COL[col]
    return f"Base!${letra}$2:${letra}${ULT}"


R_ASIG, R_AREA, R_NIVEL, R_APROBO = rango("asignatura"), rango("area"), rango("nivel_desempeno"), rango("aprobo")
R_PROM, R_NOTA1 = rango("promedio_evaluaciones"), rango("nota_primera_evaluacion")
R_ACT1, R_ACT4 = rango("activo_en_primera_evaluacion"), rango("activo_semana_4")

# quintiles de clics: se reproducen igual que en analisis_rendimiento.py (rank "first" + qcut)
ok4 = base[base["activo_semana_4"]].copy()
ok4["q"] = pd.qcut(ok4["clics_primeras_4_semanas"].rank(method="first"), 5, labels=False)
quintiles = ok4.groupby("q").agg(n=("aprobo", "size"), aprobados=("aprobo", lambda s: int((s == "Sí").sum())))
for q in range(5):
    assert int(quintiles.loc[q, "n"]) == RES["patron_clics_4_semanas"][str(q)]["n"], "n del quintil no coincide"


def estilo(celda, fuente=F_NORMAL, fmt=None, relleno=None, borde=True, alinear=None):
    celda.font = fuente
    if fmt:
        celda.number_format = fmt
    if relleno:
        celda.fill = relleno
    if borde:
        celda.border = BORDE
    if alinear:
        celda.alignment = alinear
    return celda


def encabezado(ws, fila, textos, col0=1):
    for i, t in enumerate(textos):
        estilo(ws.cell(fila, col0 + i, t), F_NEGRITA, relleno=GRIS,
               alinear=Alignment(wrap_text=True, vertical="center", horizontal="center"))


def texto(ws, ref, valor, fuente=F_NORMAL):
    ws[ref] = valor
    ws[ref].font = fuente
    return ws[ref]


wb = Workbook()

# ------------------------------------------------------------------ 1. Base
ws = wb.active
ws.title = "Base"
ws.append(list(base.columns))
columnas = []
for c in base.columns:
    s = base[c]
    columnas.append([None if pd.isna(v) else v for v in s.tolist()])
for fila in zip(*columnas):
    ws.append(fila)
for celda in ws[1]:
    celda.font = F_NEGRITA
    celda.fill = GRIS
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(base.columns))}{ULT}"
anchos = {"asignatura": 11, "area": 22, "cohorte": 15, "id_estudiante": 13, "genero": 9, "formacion_previa": 22,
          "rango_edad": 11, "intentos_previos": 9, "creditos": 9, "discapacidad": 12, "nivel_desempeno": 16,
          "aprobo": 8, "evaluaciones_entregadas": 12, "promedio_evaluaciones": 12, "nota_primera_evaluacion": 12,
          "clics_primeras_4_semanas": 12, "dia_retiro": 10, "activo_en_primera_evaluacion": 14, "activo_semana_4": 12}
for c, letra in COL.items():
    ws.column_dimensions[letra].width = anchos.get(c, 12)
ws.row_dimensions[1].height = 45
for celda in ws[1]:
    celda.alignment = Alignment(wrap_text=True, vertical="center")

# ------------------------------------------------------------------ 3. Por asignatura (antes de Indicadores en el código: el KPI la usa)
wa = wb.create_sheet("Por asignatura")
texto(wa, "A1", "Desempeño por asignatura", F_TITULO)
texto(wa, "A2", "Todas las cifras son fórmulas sobre la hoja Base. Los porcentajes se calculan sobre todas las matrículas de la asignatura.", F_NOTA)
encabezado(wa, 4, ["Asignatura", "Área", "Matrículas", "Promedio de evaluaciones"] + NIVELES + ["Aprobación"])
F0, F1 = 5, 5 + len(ASIGNATURAS) - 1  # filas 5 a 11
for i, a in enumerate(ASIGNATURAS):
    r = F0 + i
    estilo(wa.cell(r, 1, a), F_NEGRITA)
    estilo(wa.cell(r, 2, f"=INDEX({R_AREA},MATCH($A{r},{R_ASIG},0))"))
    estilo(wa.cell(r, 3, f"=COUNTIF({R_ASIG},$A{r})"), fmt=MILES)
    estilo(wa.cell(r, 4, f"=AVERAGEIFS({R_PROM},{R_ASIG},$A{r})"), fmt=UN_DEC)
    for j in range(len(NIVELES)):
        l = get_column_letter(5 + j)
        estilo(wa.cell(r, 5 + j, f"=COUNTIFS({R_ASIG},$A{r},{R_NIVEL},{l}$4)/$C{r}"), fmt=PCT)
    estilo(wa.cell(r, 9, f'=COUNTIFS({R_ASIG},$A{r},{R_APROBO},"Sí")/$C{r}'), fmt=PCT)
rt = F1 + 1  # fila 12: total
estilo(wa.cell(rt, 1, "Todas"), F_NEGRITA, relleno=CLARO)
estilo(wa.cell(rt, 2, "Todas las áreas"), relleno=CLARO)
estilo(wa.cell(rt, 3, f"=COUNTA({R_ASIG})"), F_NEGRITA, MILES, CLARO)
estilo(wa.cell(rt, 4, f"=AVERAGE({R_PROM})"), F_NEGRITA, UN_DEC, CLARO)
for j in range(len(NIVELES)):
    l = get_column_letter(5 + j)
    estilo(wa.cell(rt, 5 + j, f"=COUNTIF({R_NIVEL},{l}$4)/$C{rt}"), F_NEGRITA, PCT, CLARO)
estilo(wa.cell(rt, 9, f'=COUNTIF({R_APROBO},"Sí")/$C{rt}'), F_NEGRITA, PCT, CLARO)
texto(wa, f"A{rt + 2}", "Aprobación = (Distinción + Aprobó) / matrículas. Promedio = media de las notas de evaluación de cada matrícula; las matrículas sin notas no cuentan.", F_NOTA)
for letra, ancho in zip("ABCDEFGHI", [12, 22, 12, 14, 12, 12, 12, 12, 12]):
    wa.column_dimensions[letra].width = ancho
wa.row_dimensions[4].height = 32

g1 = BarChart()
g1.type = "col"
g1.grouping = "percentStacked"
g1.overlap = 100
g1.title = "Nivel de desempeño por asignatura"
g1.y_axis.title = "Porcentaje de matrículas"
g1.x_axis.title = "Asignatura"
g1.y_axis.number_format = "0 %"
g1.add_data(Reference(wa, min_col=5, max_col=8, min_row=4, max_row=F1), titles_from_data=True)
g1.set_categories(Reference(wa, min_col=1, min_row=F0, max_row=F1))
for serie, color in zip(g1.series, ["2E7D32", "81C784", "E57373", "9E9E9E"]):
    serie.graphicalProperties.solidFill = color
    serie.graphicalProperties.line.solidFill = color
g1.x_axis.delete = False
g1.y_axis.delete = False
g1.height, g1.width = 9, 18
wa.add_chart(g1, "K2")

g2 = BarChart()
g2.type = "col"
g2.title = "Promedio de evaluaciones por asignatura"
g2.y_axis.title = "Promedio (0 a 100)"
g2.x_axis.title = "Asignatura"
g2.add_data(Reference(wa, min_col=4, min_row=4, max_row=F1), titles_from_data=True)
g2.set_categories(Reference(wa, min_col=1, min_row=F0, max_row=F1))
g2.series[0].graphicalProperties.solidFill = "1F4E79"
g2.legend = None
g2.y_axis.scaling.min = 0
g2.y_axis.scaling.max = 100
g2.dataLabels = DataLabelList()
g2.dataLabels.showVal = True
g2.y_axis.number_format = "0"
g2.x_axis.delete = False
g2.y_axis.delete = False
g2.height, g2.width = 9, 18
wa.add_chart(g2, "K21")

# ------------------------------------------------------------------ 2. Indicadores
wi = wb.create_sheet("Indicadores", 1)
texto(wi, "A1", "Indicadores de rendimiento académico (OULAD)", F_TITULO)
estilo(wi["A3"], F_NEGRITA, relleno=GRIS)
wi["A3"] = "Asignatura"
estilo(wi["B3"], F_NEGRITA, relleno=ENTRADA, alinear=Alignment(horizontal="center"))
texto(wi, "C3", "Escriba un código (AAA a GGG) o deje la celda vacía para ver Todas.", F_NOTA)
estilo(wi["A4"], F_NEGRITA, relleno=GRIS)
wi["A4"] = "Filtro aplicado"
estilo(wi["B4"], alinear=Alignment(horizontal="center"))
wi["B4"] = '=IF($B$3="","Todas",$B$3)'
dv = DataValidation(type="list", formula1='"' + ",".join(ASIGNATURAS) + '"', allow_blank=True,
                    showErrorMessage=True, errorTitle="Asignatura", error="Use un código entre AAA y GGG, o deje la celda vacía.")
wi.add_data_validation(dv)
dv.add("B3")

CRIT = 'IF($B$3="","*",$B$3)'  # comodín: todas las asignaturas
encabezado(wi, 6, ["Indicador", "Valor", "Qué mide"])
kpis = [
    ("Matrículas", f'=IF($B$3="",COUNTA({R_ASIG}),COUNTIF({R_ASIG},$B$3))', MILES,
     "Número de matrículas (estudiante en una asignatura y cohorte) dentro del filtro."),
    ("Promedio de evaluaciones", f'=IFERROR(IF($B$3="",AVERAGE({R_PROM}),AVERAGEIFS({R_PROM},{R_ASIG},$B$3)),"")', UN_DEC,
     "Media de la nota promedio de cada matrícula (0 a 100). Las matrículas sin ninguna nota no cuentan."),
    ("Tasa de aprobación", f'=IFERROR(COUNTIFS({R_ASIG},{CRIT},{R_APROBO},"Sí")/$B$7,"")', PCT,
     "Porcentaje de matrículas que terminaron con Distinción o Aprobó."),
    ("Tasa de reprobación", f'=IFERROR(COUNTIFS({R_ASIG},{CRIT},{R_NIVEL},"Reprobó")/$B$7,"")', PCT,
     "Porcentaje de matrículas que llegaron al final y reprobaron."),
    ("Tasa de retiro", f'=IFERROR(COUNTIFS({R_ASIG},{CRIT},{R_NIVEL},"Se retiró")/$B$7,"")', PCT,
     "Porcentaje de matrículas que se retiraron antes de terminar."),
    ("Asignatura más crítica",
     f"=INDEX('Por asignatura'!$A${F0}:$A${F1},MATCH(MIN('Por asignatura'!$I${F0}:$I${F1}),'Por asignatura'!$I${F0}:$I${F1},0))",
     None, "Asignatura con la menor tasa de aprobación de las siete (no depende del filtro)."),
]
for i, (nombre, formula, fmt, nota) in enumerate(kpis):
    r = 7 + i
    estilo(wi.cell(r, 1, nombre), F_NEGRITA)
    estilo(wi.cell(r, 2, formula), F_NEGRITA, fmt, alinear=Alignment(horizontal="center"))
    estilo(wi.cell(r, 3, nota), alinear=AJUSTE)
texto(wi, "A14", "Todos los valores son fórmulas sobre la hoja Base. El desglose por asignatura está en la hoja Por asignatura.", F_NOTA)
wi.column_dimensions["A"].width = 26
wi.column_dimensions["B"].width = 16
wi.column_dimensions["C"].width = 70
for r in range(7, 13):
    wi.row_dimensions[r].height = 30

# ------------------------------------------------------------------ 4. Patrones
wp = wb.create_sheet("Patrones")
texto(wp, "A1", "Patrones tempranos de aprobación", F_TITULO)
texto(wp, "A3", "1. Aprobación según la nota de la primera evaluación (solo matrículas activas el día de la primera evaluación)", F_NEGRITA)
encabezado(wp, 4, ["Nota de la primera evaluación", "Matrículas", "Aprobados", "Tasa de aprobación"])
cond = [f'{R_NOTA1},""', f'{R_NOTA1},"<40"', f'{R_NOTA1},">=40",{R_NOTA1},"<60"',
        f'{R_NOTA1},">=60",{R_NOTA1},"<80"', f'{R_NOTA1},">=80"']
for i, (nombre, c) in enumerate(zip(RANGOS, cond)):
    r = 5 + i
    estilo(wp.cell(r, 1, nombre), F_NEGRITA)
    estilo(wp.cell(r, 2, f"=COUNTIFS({R_ACT1},TRUE,{c})"), fmt=MILES)
    estilo(wp.cell(r, 3, f'=COUNTIFS({R_ACT1},TRUE,{c},{R_APROBO},"Sí")'), fmt=MILES)
    estilo(wp.cell(r, 4, f"=C{r}/B{r}"), fmt=PCT)
estilo(wp.cell(10, 1, "Total"), F_NEGRITA, relleno=CLARO)
estilo(wp.cell(10, 2, "=SUM(B5:B9)"), F_NEGRITA, MILES, CLARO)
estilo(wp.cell(10, 3, "=SUM(C5:C9)"), F_NEGRITA, MILES, CLARO)
estilo(wp.cell(10, 4, "=C10/B10"), F_NEGRITA, PCT, CLARO)
texto(wp, "A11", "Todo con fórmulas COUNTIFS sobre la hoja Base. Se aprueba una evaluación con 40.", F_NOTA)

texto(wp, "A14", "2. Aprobación por quintil de clics en las primeras 4 semanas (solo matrículas activas al día 28)", F_NEGRITA)
encabezado(wp, 15, ["Quintil de clics", "Clics desde", "Clics hasta", "Matrículas", "Aprobados", "Tasa de aprobación"])
for q in range(5):
    r = 16 + q
    pq = RES["patron_clics_4_semanas"][str(q)]
    estilo(wp.cell(r, 1, f"Q{q + 1}"), F_NEGRITA)
    estilo(wp.cell(r, 2, pq["min"]), fmt=MILES)
    estilo(wp.cell(r, 3, pq["max"]), fmt=MILES)
    estilo(wp.cell(r, 4, int(quintiles.loc[q, "n"])), fmt=MILES)
    estilo(wp.cell(r, 5, int(quintiles.loc[q, "aprobados"])), fmt=MILES)
    estilo(wp.cell(r, 6, f"=E{r}/D{r}"), fmt=PCT)
estilo(wp.cell(21, 1, "Total"), F_NEGRITA, relleno=CLARO)
for c in "BC":
    estilo(wp[f"{c}21"], relleno=CLARO)
estilo(wp.cell(21, 4, "=SUM(D16:D20)"), F_NEGRITA, MILES, CLARO)
estilo(wp.cell(21, 5, "=SUM(E16:E20)"), F_NEGRITA, MILES, CLARO)
estilo(wp.cell(21, 6, "=E21/D21"), F_NEGRITA, PCT, CLARO)
nota_q = ("Nota: las columnas Matrículas y Aprobados son valores tomados del análisis en Python (analisis_rendimiento.py). "
          "Los quintiles se armaron por posición (rank, método first): cada uno tiene la quinta parte de las matrículas, "
          "y los estudiantes con el mismo número de clics en un borde (por ejemplo 66, 162, 295 o 543) quedan repartidos entre "
          "dos quintiles. Por eso un COUNTIFS por rango de clics no da los mismos conteos. La tasa sí es fórmula.")
c = texto(wp, "A22", nota_q, F_NOTA)
wp.merge_cells("A22:F24")
c.alignment = AJUSTE
for letra, ancho in zip("ABCDEF", [28, 13, 13, 13, 13, 14]):
    wp.column_dimensions[letra].width = ancho
wp.row_dimensions[4].height = 32
wp.row_dimensions[15].height = 32


def barras(titulo, eje_x, ref_val, ref_cat, color):
    g = BarChart()
    g.type = "col"
    g.title = titulo
    g.x_axis.title = eje_x
    g.y_axis.title = "Tasa de aprobación"
    g.add_data(ref_val, titles_from_data=True)
    g.set_categories(ref_cat)
    g.series[0].graphicalProperties.solidFill = color
    g.legend = None
    g.y_axis.scaling.min = 0
    g.y_axis.scaling.max = 1
    g.y_axis.number_format = "0 %"
    g.dataLabels = DataLabelList()
    g.dataLabels.showVal = True
    g.dataLabels.numFmt = "0.0 %"
    g.x_axis.delete = False
    g.y_axis.delete = False
    g.height, g.width = 8, 16
    return g


wp.add_chart(barras("Aprobación según la nota de la primera evaluación", "Nota de la primera evaluación",
                    Reference(wp, min_col=4, min_row=4, max_row=9), Reference(wp, min_col=1, min_row=5, max_row=9), "1F4E79"), "H2")
wp.add_chart(barras("Aprobación por quintil de clics (primeras 4 semanas)", "Quintil de clics",
                    Reference(wp, min_col=6, min_row=15, max_row=20), Reference(wp, min_col=1, min_row=16, max_row=20), "2E7D32"), "H19")

# ------------------------------------------------------------------ 5. Fuente
wf = wb.create_sheet("Fuente")
texto(wf, "A1", "Fuente, limpieza y definiciones", F_TITULO)
filas_fuente = [
    ("Fuente", "Open University Learning Analytics Dataset (OULAD). Kuzilek, J., Hlosta, M. y Zdrahal, Z. (2017). "
               "Open University Learning Analytics dataset. Scientific Data, 4, 170171."),
    ("Licencia", "CC BY 4.0"),
    ("Enlace", "https://analyse.kmi.open.ac.uk/open_dataset"),
    ("Unidad de análisis", "Una fila por matrícula: estudiante en una asignatura (AAA a GGG) y una cohorte."),
]
r = 3
for k, v in filas_fuente:
    estilo(wf.cell(r, 1, k), F_NEGRITA, relleno=GRIS)
    estilo(wf.cell(r, 2, v), alinear=AJUSTE)
    r += 1
wf["B5"].hyperlink = "https://analyse.kmi.open.ac.uk/open_dataset"

r += 1
texto(wf, f"A{r}", "Decisiones de limpieza", F_NEGRITA)
r += 1
encabezado(wf, r, ["Revisión", "Decisión", "Cantidad"])
lim = RES["limpieza"]
desc = {
    "matriculas_leidas": ("Matrículas leídas (studentInfo)", "Todas entran a la base."),
    "matriculas_duplicadas": ("Matrículas duplicadas", "No hay duplicados por estudiante, asignatura y cohorte."),
    "notas_leidas": ("Notas de evaluaciones leídas (studentAssessment)", "Se usan para el promedio y la nota de la primera evaluación."),
    "notas_sin_puntaje": ("Notas sin puntaje", "Se descartan: no cuentan en el promedio ni como entregadas."),
    "notas_fuera_de_rango": ("Notas fuera de 0 a 100", "No hay; no fue necesario corregir."),
    "notas_trasladadas_de_intento_anterior": ("Notas trasladadas de un intento anterior", "Se conservan: son notas válidas del estudiante."),
    "examenes_sin_fecha": ("Exámenes sin fecha", "No afectan: la primera evaluación se toma entre las evaluaciones del semestre, sin el examen."),
    "imd_vacio": ("Índice socioeconómico (IMD) vacío", "Sin vacíos; la variable no se usa en este taller."),
}
for k, v in lim.items():
    r += 1
    nombre, decision = desc.get(k, (k, ""))
    estilo(wf.cell(r, 1, nombre), alinear=AJUSTE)
    estilo(wf.cell(r, 2, decision), alinear=AJUSTE)
    estilo(wf.cell(r, 3, v), fmt=MILES)

r += 2
texto(wf, f"A{r}", "Definiciones", F_NEGRITA)
r += 1
encabezado(wf, r, ["Término", "Definición"])
defs = [
    ("Aprobación", "(Distinción + Aprobó) / todas las matrículas. Los retirados cuentan en el denominador."),
    ("Reprobación", "Matrículas con resultado final Reprobó / todas las matrículas."),
    ("Retiro", "Matrículas con resultado final Se retiró / todas las matrículas."),
    ("Promedio de evaluaciones", "Media de las notas (0 a 100) que entregó cada matrícula; el indicador es la media de esos promedios. Las matrículas sin notas quedan vacías y no cuentan."),
    ("Aprobar una evaluación", "Nota de 40 o más."),
    ("Primera evaluación", "La primera evaluación del semestre (sin contar el examen final) de cada asignatura y cohorte."),
    ("Activo en la primera evaluación", "Seguía inscrito el día de la primera evaluación (no se retiró antes)."),
    ("Activo en la semana 4", "Seguía inscrito al día 28 del curso."),
    ("Clics en las primeras 4 semanas", "Suma de clics en la plataforma virtual antes del día 28 del curso (incluye los días previos al inicio)."),
    ("Asignatura más crítica", "La asignatura con la menor tasa de aprobación."),
]
for k, v in defs:
    r += 1
    estilo(wf.cell(r, 1, k), F_NEGRITA, alinear=AJUSTE)
    estilo(wf.cell(r, 2, v), alinear=AJUSTE)
wf.column_dimensions["A"].width = 34
wf.column_dimensions["B"].width = 80
wf.column_dimensions["C"].width = 12

# ------------------------------------------------------------------ orden, metadatos y guardado
wb._sheets = [wb["Base"], wb["Indicadores"], wb["Por asignatura"], wb["Patrones"], wb["Fuente"]]
wb.active = 1
for hoja in wb.worksheets:
    hoja.sheet_view.tabSelected = hoja.title == "Indicadores"
wb.properties.creator = ""
wb.properties.lastModifiedBy = ""
wb.properties.title = None
wb.calculation.fullCalcOnLoad = True
SALIDA.parent.mkdir(exist_ok=True)
wb.save(SALIDA)
print("Guardado:", SALIDA)
