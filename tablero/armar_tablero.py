"""Junta index.html, chart.umd.js y datos.js en un solo archivo que abre sin internet."""
from pathlib import Path

T = Path(__file__).resolve().parent
html = (T / "index.html").read_text(encoding="utf-8")
for nombre in ("chart.umd.js", "xlsx.full.min.js", "datos.js"):
    codigo = (T / nombre).read_text(encoding="utf-8").replace("</script", "<\\/script")
    html = html.replace(f'<script src="{nombre}"></script>', f"<script>\n{codigo}\n</script>")
salida = T.parent / "entrega" / "Tablero-Rendimiento-Academico.html"
salida.write_text(html, encoding="utf-8")
print(salida, salida.stat().st_size // 1024, "KB")
