"""One-time script: convert the highlighted RUIZ DELGADO example letter into a
reusable docxtpl template by replacing each yellow-highlighted run-group with a
{tag}, and wrapping CLAUSULA QUINTA (Garantizado) in a {#garantizado}...{/garantizado}
conditional block. Preserves all non-highlighted formatting (bold headers, tabs, etc.)
because only the highlighted runs are touched.
"""
import docx
from docx.enum.text import WD_COLOR_INDEX

SRC = "template.docx"
OUT = "plantilla_otrosi_cargo.docx"

# Exact literal text -> jinja tag (checked via grouping script, order matters:
# longer/more specific matches first is not needed since these are exact full-group matches)
REPLACEMENTS = {
    "CERON HOYOS JUAN SEBASTIAN": "{nombre}",
    "1.015.441.511": "{cedula}",
    "Analista Nacional TAT": "{cargo_actual}",
    "Analista Conciliación Claro Pay": "{cargo_nuevo}",
    "Dirección Canales de Venta Agentes": "{gerencia_nueva}",
    "Bogotá D.C": "{ciudad}",
    # date spans (day is fixed at 1, only month/year vary per confirmed business rule)
    "1 de octubre de 2026": "1 de {mes} de {anio}",
    "1 de octubre del año 2026": "1 de {mes} del año {anio}",
    "(1) de Octubre de 2026": "(1) de {Mes} de {anio}",
    "01° de  Octubre de 2026": "01° de {Mes} de {anio}",
    "01° de Octubre de 2026": "01° de {Mes} de {anio}",
}

CLAUSULA_QUINTA_START = "CLÁUSULA QUINTA"


def replace_highlighted_groups(paragraph):
    runs = paragraph.runs
    i = 0
    while i < len(runs):
        if runs[i].font.highlight_color == WD_COLOR_INDEX.YELLOW:
            j = i
            group_runs = []
            while j < len(runs) and runs[j].font.highlight_color == WD_COLOR_INDEX.YELLOW:
                group_runs.append(runs[j])
                j += 1
            combined = "".join(r.text for r in group_runs)
            if combined not in REPLACEMENTS:
                raise ValueError(f"Unmapped highlighted span: {combined!r} in paragraph: {paragraph.text!r}")
            group_runs[0].text = REPLACEMENTS[combined]
            group_runs[0].font.highlight_color = None
            for r in group_runs[1:]:
                r.text = ""
            i = j
        else:
            i += 1


def wrap_garantizado(doc):
    for p in doc.paragraphs:
        if p.text.strip().startswith(CLAUSULA_QUINTA_START):
            # put the opening tag in front of the first run's text and the
            # closing tag at the end of the last run's text (same paragraph)
            runs = p.runs
            runs[0].text = "{#garantizado}" + runs[0].text
            runs[-1].text = runs[-1].text + "{/garantizado}"
            return
    raise ValueError("CLAUSULA QUINTA paragraph not found")


def main():
    doc = docx.Document(SRC)
    for p in doc.paragraphs:
        replace_highlighted_groups(p)
    wrap_garantizado(doc)
    doc.save(OUT)
    print("Saved", OUT)


if __name__ == "__main__":
    main()
