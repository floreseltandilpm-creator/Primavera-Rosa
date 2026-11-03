"""
Clasifica los análisis foliares (PDF del laboratorio GR Chía) por válvula.

Uso:  python clasificar_foliares.py        (dentro de la carpeta Analisis_Foliar)
Requiere:  pip install pypdf

Lee todos los PDF de esta carpeta, extrae sector, bloques, válvulas, semana y
fechas, y genera indice_foliar.js. El tablero (Index.html) usa ese índice para
que el botón "Último Análisis Foliar" de cada válvula abra el PDF más reciente.
No hace falta renombrar los PDF: un mismo informe puede cubrir varias válvulas.
"""
import json, re, sys
from pathlib import Path
from pypdf import PdfReader

CARPETA = Path(__file__).resolve().parent
FECHA = r"\d{4}-\d{2}-\d{2}"
FILA = re.compile(
    rf"(?P<solicitud>\d{{5,}})\s+(?P<semana>\d{{1,2}})\s+(?P<muestreo>{FECHA})\s+"
    rf"(?P<recepcion>{FECHA})\s+(?P<informe>{FECHA})\s+(?P<parte>.+?)\s+"
    r"(?P<sector>\d+)\s+(?P<bloques>[A-Za-z0-9,\s]+?)\s+(?P<valvulas>\d[\d,\s\-y]*)$",
    re.M)

def expandir_valvulas(txt):
    vals = set()
    for parte in re.split(r"[,\sy]+", txt.strip()):
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-", 1)
            if a.isdigit() and b.isdigit():
                vals.update(range(int(a), int(b) + 1))
        elif parte.isdigit():
            vals.add(int(parte))
    return sorted(vals)

def leer(pdf):
    texto = PdfReader(pdf).pages[0].extract_text() or ""
    m = FILA.search(texto)
    if not m:
        return None
    lab = re.search(r"\b(TV\d+)\b", texto)
    # Variedad = todo lo que va entre "Rose - Color" y el estado fenológico (puede tener varias palabras)
    var = re.search(r"Rose\s*-\s*\S+\s+(?P<variedad>.+?)\s+(?P<estado>Productivo|Vegetativo|Improductivo|Pinch|Poda|Corte|Renovaci[oó]n|Floraci[oó]n)\s+(?P<sint>.+)", texto)
    return {
        "archivo": pdf.name,
        "laboratorio": lab.group(1) if lab else "",
        "solicitud": m["solicitud"],
        "semana": int(m["semana"]),
        "fechaMuestreo": m["muestreo"],
        "fechaInforme": m["informe"],
        "parte": m["parte"].strip(),
        "sector": int(m["sector"]),
        "bloques": [b.strip() for b in m["bloques"].split(",") if b.strip()],
        "valvulas": expandir_valvulas(m["valvulas"]),
        "variedad": var["variedad"].strip() if var else "",
        "estado": var["estado"] if var else "",
        "sintomatologia": var["sint"].strip() if var else "",
    }

def main():
    registros, fallidos = [], []
    for pdf in sorted(CARPETA.glob("*.pdf")):
        try:
            r = leer(pdf)
        except Exception as e:
            r = None
            print(f"  ! Error leyendo {pdf.name}: {e}")
        (registros.append(r) if r else fallidos.append(pdf.name))
    registros.sort(key=lambda r: (r["fechaMuestreo"], r["fechaInforme"]))
    js = ("/* Generado por clasificar_foliares.py — no editar a mano. */\n"
          "window.ANALISIS_FOLIAR = " + json.dumps(registros, ensure_ascii=False, indent=1) + ";\n")
    (CARPETA / "indice_foliar.js").write_text(js, encoding="utf-8")
    for r in registros:
        print(f"  ✓ {r['archivo']}: sector {r['sector']}, bloque {', '.join(r['bloques'])}, "
              f"válvulas {', '.join(map(str, r['valvulas']))} (muestreo {r['fechaMuestreo']})")
    for f in fallidos:
        print(f"  ✗ {f}: no se reconoció el formato del informe")
    print(f"\n{len(registros)} análisis clasificados → indice_foliar.js")
    return 1 if fallidos else 0

if __name__ == "__main__":
    sys.exit(main())
