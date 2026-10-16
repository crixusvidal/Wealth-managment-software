"""Actualiza la carpeta del informe consolidado con los ÚLTIMOS archivos finales y elimina lo viejo.

Origen:  secciones/informe_final/documento/Mallplaza_Informe_CFA_{ES,EN}.{pdf,docx}  (los escribe maquetar.py y typst/armar.py)
Destino: informe_consolidado/vigente/   ← la versión actual
         informe_consolidado/anterior/  ← la penúltima (regla del proyecto: solo dos versiones)

Si un archivo cambió (por hash) respecto de `vigente/`, el de `vigente/` pasa a `anterior/` (reemplazando lo que hubiera ahí);
lo que no cambió no se toca. Cualquier otro archivo en esas carpetas se borra. Escribe README.md con fecha, páginas y hashes.

Uso (desde la raíz del repo):   python informe_consolidado/actualizar.py
"""
import hashlib
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
ORIGEN = RAIZ / "secciones" / "informe_final" / "documento"
AQUI = Path(__file__).parent
VIGENTE, ANTERIOR = AQUI / "vigente", AQUI / "anterior"
NOMBRES = [f"Mallplaza_Informe_CFA_{l}.{e}" for l in ("ES", "EN") for e in ("pdf", "docx")]


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()[:12]


def paginas(p):
    out = subprocess.run(["pdfinfo", str(p)], capture_output=True, text=True).stdout
    m = re.search(r"Pages:\s+(\d+)", out)
    return int(m.group(1)) if m else "—"


def main():
    VIGENTE.mkdir(exist_ok=True)
    ANTERIOR.mkdir(exist_ok=True)
    cambios = []
    for n in NOMBRES:
        src = ORIGEN / n
        if not src.exists():
            print(f"AVISO: falta {src.relative_to(RAIZ)}")
            continue
        dst = VIGENTE / n
        if dst.exists() and sha(dst) == sha(src):
            continue
        if dst.exists():                              # la vigente pasa a anterior (reemplaza a la penúltima)
            shutil.copy2(dst, ANTERIOR / n)
        shutil.copy2(src, dst)
        cambios.append(n)
    for carpeta in (VIGENTE, ANTERIOR):               # borra todo lo que no sea parte del informe
        for p in carpeta.iterdir():
            if p.name not in NOMBRES and p.name != ".gitkeep":
                p.unlink()
    filas = []
    for n in NOMBRES:
        p = VIGENTE / n
        if p.exists():
            filas.append(f"| `vigente/{n}` | {datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M} | "
                         f"{paginas(p) if n.endswith('pdf') else '—'} | `{sha(p)}` | {'actualizado ahora' if n in cambios else 'sin cambios'} |")
    (AQUI / "README.md").write_text("\n".join([
        "# Informe consolidado (Mallplaza, CFA Institute Research Challenge)", "",
        "Entregables finales en español e inglés. **Entregar los PDF** (cuerpo maquetado con Typst + anexos); los Word sirven para editar.", "",
        "| Archivo | Fecha | Páginas | Hash | Estado |", "|---|---|---|---|---|", *filas, "",
        "- `vigente/`: la versión actual. `anterior/`: la penúltima (regla del proyecto: solo dos versiones).",
        "- Se actualiza con `python informe_consolidado/actualizar.py` (lo hace también `generar_todo.sh`); no editar a mano.",
        "- Cómo se arma: `secciones/informe_final/` (Word: `maquetar.py`; PDF: `typst/armar.py`). Mapa de carpetas en `narrativa_y_coherencia.md`.",
        "- Cifras clave vigentes para el LLM: `contexto_llm/cifras_clave.md`.", ""]), encoding="utf-8")
    print("actualizados:", cambios or "ninguno (todo igual)")


if __name__ == "__main__":
    main()
