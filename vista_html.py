"""Vista HTML del informe (para ver en el navegador y descargar como .html)."""
import html
import re
from pathlib import Path

BASE = Path(__file__).parent.resolve()


def _fmt(texto):
    t = html.escape(texto or "")
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"\[\[(.+?)\]\]", r"<mark>\1</mark>", t)
    return t


def _fig(prefijo, ruta, epigrafe, ancho):
    src = html.escape(prefijo + ruta)
    cap = f"<figcaption>{_fmt(epigrafe)}</figcaption>" if epigrafe else ""
    return f'<figure><img src="{src}" alt="{html.escape(epigrafe or "Imagen")}" style="max-width:{int(ancho * 38)}px">{cap}</figure>'


def spec_a_html(spec, prefijo=""):
    o = []
    o.append(
        '<table class="cabecera"><tr><td rowspan="4" class="logo"><b>ENFOQUE</b><br>INGENIERÍA</td>'
        f'<td rowspan="2" class="tit">{_fmt(spec.get("tipo_label", "INFORME TÉCNICO"))}</td><td>CÓDIGO {_fmt(spec.get("codigo", "DOC"))}</td></tr>'
        f'<tr><td>REVISIÓN {_fmt(spec.get("revision", "00"))}</td></tr>'
        f'<tr><td rowspan="2" class="tit">{_fmt(spec.get("subtitulo", ""))}</td><td>{_fmt(spec.get("fecha", ""))}</td></tr>'
        "<tr><td>&nbsp;</td></tr></table>"
    )
    o.append(f'<h1 class="titulo">{_fmt((spec.get("titulo") or "").upper())}</h1>')
    for n, s in enumerate(spec.get("secciones", []), 1):
        o.append(f"<h2>{n}. {_fmt(s.get('titulo', ''))}</h2>")
        for b in s.get("bloques", []):
            t = b.get("type")
            if t == "p":
                o.append(f"<p>{_fmt(b.get('text', ''))}</p>")
            elif t == "bullets":
                o.append("<ul>" + "".join(f"<li>{_fmt(i)}</li>" for i in b.get("items", [])) + "</ul>")
            elif t == "numbered":
                o.append("<ol>" + "".join(f"<li>{_fmt(i)}</li>" for i in b.get("items", [])) + "</ol>")
            elif t == "table":
                th = "".join(f"<th>{_fmt(str(c))}</th>" for c in b.get("header", []))
                filas = "".join("<tr>" + "".join(f"<td>{_fmt(str(c if c is not None else ''))}</td>" for c in r) + "</tr>" for r in b.get("rows", []))
                o.append(f'<div class="scroll"><table class="datos"><thead><tr>{th}</tr></thead><tbody>{filas}</tbody></table></div>')
            elif t == "figure":
                o.append(_fig(prefijo, b.get("path", ""), b.get("caption"), b.get("widthCm", 14)))
    anexo = spec.get("anexo")
    if anexo and anexo.get("fotos"):
        o.append(f'<h2 class="anexo">ANEXO: {_fmt(anexo.get("titulo", "Registro fotográfico"))}</h2>')
        for i, f in enumerate(anexo["fotos"], 1):
            o.append(_fig(prefijo, f.get("path", ""), f.get("caption") or f"Foto {i}", f.get("widthCm", 11)))
    o.append('<table class="firmas"><tr><td>Elaboró:</td><td>Revisó:</td><td>Aprobó:</td></tr></table>')
    return "\n".join(o)


def pagina_html(cuerpo, titulo="Informe"):
    css = (BASE / "static" / "informe.css").read_text(encoding="utf-8")
    return (
        '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        f"<title>{html.escape(titulo)}</title><style>{css}</style></head>"
        f'<body><div class="informe">{cuerpo}</div></body></html>'
    )
