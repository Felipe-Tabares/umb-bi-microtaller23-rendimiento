# Microtaller 23 · Análisis de rendimiento académico

Inteligencia de Negocios, Universidad Manuela Beltrán. Andrés Felipe Tabares Benavides, noviembre de 2026.

Identifica patrones en los resultados académicos de 32.593 matrículas reales y anónimas de la Open University (7 asignaturas, 2013 y 2014) y los muestra en un tablero interactivo con promedios, asignaturas críticas y niveles de desempeño.

**Tablero en línea:** https://felipe-tabares.github.io/umb-bi-microtaller23-rendimiento/tablero/

## Cargar sus propios datos

El tablero no sirve solo para la base incluida. Con el botón **Cargar Excel o CSV** (o arrastrando el archivo a la página) se sube una hoja de notas y todo se recalcula en el navegador; el archivo no sale del computador.

- Reconoce las columnas por el nombre: `asignatura` (o materia), `nota_1`, `nota_2`... (o `corte_1`, `parcial_1`), `nota_final` (o definitiva), `resultado`, `periodo`, `jornada` y `actividad_primer_mes`.
- Detecta la escala (0 a 5, 0 a 10 o 0 a 100) y deja cambiar la nota mínima para aprobar.
- Si no hay columna de resultado, el nivel sale de la nota: distinción desde el 90 % de la escala, aprobó desde la nota mínima y sin nota cuenta como retiro.
- **Descargar plantilla** baja un Excel con las columnas y las instrucciones.
- `entrega/Ejemplo-FICTICIO-notas-0-a-5.xlsx` es un archivo **ficticio** (`analisis/ejemplo_ficticio.py`) para probar la carga.

## Qué hay aquí

| Ruta | Qué es |
|---|---|
| `analisis/analisis_rendimiento.py` | Todo el proceso: lee, limpia, cruza, arma la base y calcula los indicadores |
| `analisis/resultados.json` | Las cifras que salen del análisis |
| `datos/base_matriculas.csv` | La base construida: una fila por estudiante y asignatura |
| `datos/rendimiento.db` | La misma base en SQLite (tablas `matricula`, `evaluacion`, `nota`) |
| `tablero/` | El tablero (`index.html` + `datos.js` + Chart.js + SheetJS para leer Excel) |
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
