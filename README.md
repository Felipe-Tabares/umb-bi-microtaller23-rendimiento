# Microtaller 23 · Análisis de rendimiento académico

Inteligencia de Negocios, Universidad Manuela Beltrán. Andrés Felipe Tabares Benavides, noviembre de 2026.

Identifica patrones en los resultados académicos de 32.593 matrículas reales y anónimas de la Open University (7 asignaturas, 2013 y 2014) y los muestra en un tablero interactivo con promedios, asignaturas críticas y niveles de desempeño.

## Qué hay aquí

| Ruta | Qué es |
|---|---|
| `analisis/analisis_rendimiento.py` | Todo el proceso: lee, limpia, cruza, arma la base y calcula los indicadores |
| `analisis/resultados.json` | Las cifras que salen del análisis |
| `datos/base_matriculas.csv` | La base construida: una fila por estudiante y asignatura |
| `datos/rendimiento.db` | La misma base en SQLite (tablas `matricula`, `evaluacion`, `nota`) |
| `tablero/` | El tablero (`index.html` + `datos.js` + Chart.js) |
| `entrega/Tablero-Rendimiento-Academico.html` | El tablero en un solo archivo; abre con doble clic, sin internet |
| `entrega/*.xlsx` | La base y los indicadores con fórmulas de Excel |

## Cómo reproducirlo

1. Bajar la base original (unos 47 MB comprimida) de https://archive.ics.uci.edu/dataset/349/open+university+learning+analytics+dataset y descomprimirla en `datos/oulad/`.
2. `pip install -r requirements.txt`
3. `python analisis/analisis_rendimiento.py` (unos 4 segundos).
4. `python tablero/armar_tablero.py` para volver a armar el tablero de un solo archivo.

La base original no está en el repositorio porque el archivo de clics pesa 453 MB.

## Fuente

Kuzilek, J., Hlosta, M. y Zdrahal, Z. (2017). Open University Learning Analytics dataset. *Scientific Data, 4*, 170171. https://doi.org/10.1038/sdata.2017.171. Licencia CC BY 4.0.
