"""
Bot de informes – Enfoque Ingeniería
Flask + Gemini (gratis con clave de Google AI Studio).
Recibe datos, imágenes y audios desde una página HTML y genera el informe en Word y HTML.
"""
import io
import json
import os
import re
import time
import uuid
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_from_directory
from PIL import Image, ImageOps

import base64

from gemini_rest import ErrorGemini, generar_texto
from generador_docx import construir_docx
from vista_html import pagina_html, spec_a_html
from modulo_db import crear_tablas, guardar_informe
BASE = Path(__file__).parent.resolve()
load_dotenv(BASE / ".env", override=True)

API_KEY = os.getenv("GEMINI_API_KEY", "").strip().strip('"').strip("'")
MODELO = os.getenv("GEMINI_MODEL", "gemini-flash-latest").strip()
SALIDA = BASE / "salida"
SALIDA.mkdir(exist_ok=True)
LOGO = BASE / "static" / "logo.png"

TIPOS = {
    "avance": "Informe de avance de obra",
    "metodo-constructivo": "Método constructivo",
    "elementos-necesarios": "Elementos necesarios",
    "procedimiento": "Procedimiento de trabajo",
}
AUDIO_MIME = {
    ".mp3": "audio/mp3", ".wav": "audio/wav", ".ogg": "audio/ogg", ".opus": "audio/ogg", ".oga": "audio/ogg",
    ".flac": "audio/flac", ".aac": "audio/aac", ".m4a": "audio/aac", ".aiff": "audio/aiff", ".aif": "audio/aiff",
}
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp"}
MAX_FOTOS = 30
MAX_BYTES_IA = 18 * 1024 * 1024   # tope de lo que se envía a la IA en una sola consulta

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024


class ErrorBot(Exception):
    """Error con mensaje claro para mostrar al usuario."""


# --------------------------------------------------------------------------- Gemini
def estado_clave() -> str:
    if not API_KEY or "PEGA" in API_KEY.upper():
        return "FALTA: no hay clave en el archivo .env"
    return f"cargada ({API_KEY[:4]}…{API_KEY[-4:]}, {len(API_KEY)} caracteres)"


def comprobar_clave():
    if not API_KEY or "PEGA" in API_KEY.upper():
        raise ErrorBot("No hay clave de Gemini cargada. Creá el archivo .env junto a app.py con GEMINI_API_KEY=tu_clave (ver README).")


REGLAS = """Sos el redactor técnico de Enfoque Ingeniería, una oficina técnica de ingeniería y construcción en Argentina.
Redactás informes técnicos de obra en español rioplatense neutro, en tercera persona y con tono técnico, sin adjetivos innecesarios.

REGLAS:
- Cada sección empieza por lo que se hizo o se propone, no por generalidades.
- Todo dato que dio el usuario va como texto normal.
- Todo lo que redactes por tu cuenta y requiera verificación (normas, rendimientos, cantidades estimadas, riesgos, pasos, secuencias) va entre [[doble corchete]]; el programa lo resalta en amarillo para que el usuario lo revise.
- Si falta un dato esencial, escribí [[completar]]. Nunca inventes cantidades, rendimientos ni resultados de cálculo.
- Citá normativa (CIRSOC, Ley 19.587, Decreto 911/96, etc.) solo si la dio el usuario o marcala con [[ ]].
- Los audios son notas de voz del usuario: escuchalos e incorporá su contenido (datos, observaciones, instrucciones) al informe.
- Las imágenes son fotos de obra: describí lo que se ve de forma técnica y prudente, sin afirmar lo que no se distingue.

FORMATO DE RESPUESTA (obligatorio):
- Respondé ÚNICAMENTE con un objeto JSON válido con la misma estructura de la plantilla recibida:
  tipo_label, subtitulo, codigo, revision, fecha, titulo, secciones, anexo.
- secciones = lista de {"titulo": "...", "bloques": [...]}. No numeres los títulos (se numeran solos).
- Bloques permitidos: {"type":"p","text":"..."}, {"type":"bullets","items":[...]}, {"type":"numbered","items":[...]},
  {"type":"table","header":[...],"rows":[[...]],"widths":[...]} (widths suman 8504), {"type":"figure","path":"...","caption":"...","widthCm":14}.
- Texto admite **negrita** y [[a verificar]].
- Conservá las secciones de la plantilla; podés agregar o quitar bloques según el contenido.
- Incluí TODAS las fotos en anexo.fotos con exactamente el 'path' indicado y un epígrafe 'Foto N: descripción – fecha (si se conoce)'.
- No agregues cuadro de firmas."""


def llamar_gemini(partes, sistema=REGLAS):
    comprobar_clave()
    for intento in (1, 2):
        try:
            texto = generar_texto(API_KEY, MODELO, partes, sistema=sistema)
            return _json(texto)
        except ErrorGemini as e:
            if e.clave_invalida:
                raise ErrorBot(
                    "Google rechazó la clave (" + e.mensaje + "). Ejecutá probar_clave.py para diagnosticar. "
                    "Si la clave se creó en AI Studio (empieza con AQ.), verificá que esté completa en el .env y que el proyecto tenga la API de Gemini habilitada."
                )
            if e.codigo == 429:
                raise ErrorBot("Se alcanzó el límite de uso de Gemini. Esperá un minuto y probá de nuevo (o reducí fotos/audios).")
            if e.codigo == 404:
                raise ErrorBot(f"El modelo '{MODELO}' no existe para tu clave. Ejecutá probar_clave.py para ver los disponibles y ajustá GEMINI_MODEL en .env.")
            raise ErrorBot(f"Error de Gemini ({e.codigo}): {e.mensaje}")
        except (ValueError, json.JSONDecodeError):
            if intento == 2:
                raise ErrorBot("La IA devolvió una respuesta que no se pudo leer. Probá de nuevo.")
    raise ErrorBot("No se pudo obtener respuesta de la IA.")


def _adjunto(datos: bytes, mime: str) -> dict:
    return {"inline_data": {"mime_type": mime, "data": base64.standard_b64encode(datos).decode()}}


def _json(texto: str) -> dict:
    texto = re.sub(r"^```(?:json)?|```$", "", texto.strip(), flags=re.M).strip()
    ini, fin = texto.find("{"), texto.rfind("}")
    if ini < 0 or fin < 0:
        raise ValueError("sin JSON")
    return json.loads(texto[ini:fin + 1])


# --------------------------------------------------------------------------- archivos
def guardar_imagenes(archivos, carpeta: Path):
    fotos = []
    validos = [f for f in archivos if f and f.filename and Path(f.filename).suffix.lower() in IMG_EXT]
    for i, f in enumerate(validos[:MAX_FOTOS], 1):
        with Image.open(f.stream) as im:
            fecha = None
            try:
                ex = im.getexif()
                v = ex.get_ifd(0x8769).get(36867) or ex.get(306)
                if v:
                    fecha = datetime.strptime(str(v), "%Y:%m:%d %H:%M:%S").strftime("%d/%m/%Y")
            except Exception:
                pass
            im = ImageOps.exif_transpose(im).convert("RGB")
            im.thumbnail((1800, 1800))
            nombre = f"img_{i:02d}.jpg"
            im.save(carpeta / nombre, "JPEG", quality=85)
            ia = im.copy()
        ia.thumbnail((1280, 1280))
        buf = io.BytesIO()
        ia.save(buf, "JPEG", quality=80)
        fotos.append({"rel": nombre, "original": f.filename, "fecha": fecha, "bytes": buf.getvalue()})
    omitidas = len(validos) - len(fotos)
    return fotos, omitidas


def leer_audios(archivos):
    audios = []
    for f in archivos:
        if not f or not f.filename:
            continue
        ext = Path(f.filename).suffix.lower()
        if ext not in AUDIO_MIME:
            raise ErrorBot(f"Formato de audio no soportado: {f.filename}. Usá mp3, wav, ogg/opus, m4a, aac o flac.")
        audios.append({"nombre": f.filename, "mime": AUDIO_MIME[ext], "bytes": f.read()})
    return audios


def asegurar_fotos(spec: dict, nombres: list):
    usadas = set()
    for s in spec.get("secciones", []):
        for b in s.get("bloques", []):
            if isinstance(b, dict) and b.get("type") == "figure":
                usadas.add(b.get("path"))
    anexo = spec.get("anexo")
    if not isinstance(anexo, dict):
        anexo = spec["anexo"] = {"titulo": "Registro fotográfico", "fotos": []}
    anexo.setdefault("fotos", [])
    usadas |= {f.get("path") for f in anexo["fotos"] if isinstance(f, dict)}
    for i, n in enumerate(nombres, 1):
        if n not in usadas:
            anexo["fotos"].append({"path": n, "caption": f"Foto {i}: [[completar descripción]]"})


def limpiar_spec(spec: dict) -> dict:
    if not isinstance(spec, dict) or not isinstance(spec.get("secciones"), list):
        raise ErrorBot("La IA no devolvió la estructura esperada del informe. Probá de nuevo.")
    permitidos = {"p", "bullets", "numbered", "table", "figure"}
    secciones = []
    for s in spec["secciones"]:
        if not isinstance(s, dict):
            continue
        s["bloques"] = [b for b in s.get("bloques", []) if isinstance(b, dict) and b.get("type") in permitidos]
        secciones.append(s)
    spec["secciones"] = secciones
    return spec


def guardar_resultado(spec: dict, carpeta: Path, id_: str):
    (carpeta / "informe.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    construir_docx(spec, carpeta, carpeta / "informe.docx", logo=LOGO)
    cuerpo = spec_a_html(spec, prefijo="")
    (carpeta / "informe.html").write_text(pagina_html(cuerpo, spec.get("titulo", "Informe")), encoding="utf-8")
    return {
        "ok": True,
        "id": id_,
        "html": spec_a_html(spec, prefijo=f"/salida/{id_}/"),
        "docx": f"/salida/{id_}/informe.docx",
        "htmlfile": f"/salida/{id_}/informe.html",
    }


# --------------------------------------------------------------------------- rutas
@app.get("/")
def index():
    return render_template("index.html", tipos=TIPOS, hoy=datetime.now().strftime("%d/%m/%Y"), modelo=MODELO)


@app.post("/generar")
def generar():
    try:
        tipo = request.form.get("tipo", "avance")
        if tipo not in TIPOS:
            raise ErrorBot("Tipo de informe inválido.")
        obra = request.form.get("obra", "").strip()
        codigo = request.form.get("codigo", "").strip() or "[código]"
        revision = request.form.get("revision", "").strip() or "00"
        fecha = request.form.get("fecha", "").strip() or datetime.now().strftime("%d/%m/%Y")
        datos = request.form.get("datos", "").strip()
        sem_desde = request.form.get("semana_desde", "").strip()
        sem_hasta = request.form.get("semana_hasta", "").strip()

        periodo = ""
        if sem_desde or sem_hasta:
            periodo = (
                f"PERÍODO INFORMADO: desde {sem_desde or '[[completar]]'} hasta {sem_hasta or '[[completar]]'} "
                "(usalo en el título y en la sección Objeto del informe)\n"
            )

        modo_prueba = request.form.get("modo_prueba") == "on"

        id_ = f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:4]}"
        carpeta = SALIDA / id_
        carpeta.mkdir()

        fotos, omitidas = guardar_imagenes(request.files.getlist("imagenes"), carpeta)
        audios = [] if modo_prueba else leer_audios(request.files.getlist("audios"))

        if not (datos or fotos or audios):
            raise ErrorBot("Cargá al menos algún dato, una imagen o un audio.")

        plantilla = json.loads((BASE / "plantillas" / f"{tipo}.json").read_text(encoding="utf-8"))

        if modo_prueba:
            spec = plantilla
        else:
            texto = (
                f"TIPO DE INFORME: {tipo} ({TIPOS[tipo]})\nOBRA / PROYECTO: {obra or '[[completar]]'}\n"
                f"CÓDIGO: {codigo}\nREVISIÓN: {revision}\nFECHA: {fecha}\n{periodo}\n"
                f"DATOS DEL USUARIO:\n{datos or '(sin texto; usá las imágenes y los audios)'}\n\n"
                f"PLANTILLA A COMPLETAR (JSON):\n{json.dumps(plantilla, ensure_ascii=False, indent=1)}\n\n"
                f"Cantidad de fotos: {len(fotos)} | Cantidad de audios: {len(audios)}"
            )

            partes = [{"text": texto}]
            total = len(texto)

            for i, f in enumerate(fotos, 1):
                etiqueta = f"Foto {i} – path: {f['rel']} – archivo original: {f['original']}"
                if f["fecha"]:
                    etiqueta += f" – fecha de toma: {f['fecha']}"
                partes += [{"text": etiqueta}, _adjunto(f["bytes"], "image/jpeg")]
                total += len(f["bytes"])

            for a in audios:
                partes += [{"text": f"Nota de voz: {a['nombre']}"}, _adjunto(a["bytes"], a["mime"])]
                total += len(a["bytes"])

            if total > MAX_BYTES_IA:
                raise ErrorBot("Hay demasiado material para una sola consulta (límite ~18 MB entre fotos y audios). Usá menos audios o audios más cortos.")

            spec = llamar_gemini(partes)

        # ------------------ Procesamiento final del informe ------------------
        spec = limpiar_spec(spec)
        spec["codigo"], spec["revision"], spec["fecha"] = codigo, revision, fecha
        spec.setdefault("tipo_label", "INFORME TÉCNICO")

        if obra and not modo_prueba:
            spec.setdefault("titulo", obra)

        asegurar_fotos(spec, [f["rel"] for f in fotos])

        res = guardar_resultado(spec, carpeta, id_)
        
        res["avisos"] = [f"Se usaron las primeras {MAX_FOTOS} fotos; se omitieron {omitidas}."] if omitidas else []
        return jsonify(res)

    except ErrorBot as e:
        return jsonify(ok=False, error=str(e)), 400

    except Exception as e:  # error inesperado
        app.logger.exception("Error inesperado")
        return jsonify(ok=False, error=f"Error inesperado: {e}"), 500



@app.post("/ajustar")
def ajustar():
    try:
        datos = request.get_json(force=True)
        id_ = str(datos.get("id", ""))
        instruccion = str(datos.get("instruccion", "")).strip()
        if not re.fullmatch(r"[\w-]+", id_) or not instruccion:
            raise ErrorBot("Falta el informe o la instrucción.")
        carpeta = SALIDA / id_
        ruta = carpeta / "informe.json"
        if not ruta.exists():
            raise ErrorBot("No se encontró el informe a ajustar.")
        actual = json.loads(ruta.read_text(encoding="utf-8"))
        partes = [{"text":
            "INFORME ACTUAL (JSON):\n" + json.dumps(actual, ensure_ascii=False) +
            "\n\nINSTRUCCIÓN DEL USUARIO PARA MODIFICARLO:\n" + instruccion +
            "\n\nDevolvé el JSON COMPLETO del informe con los cambios aplicados. Conservá los 'path' de las fotos y todo lo que no se pida cambiar."
        }]
        spec = limpiar_spec(llamar_gemini(partes))
        for k in ("codigo", "revision", "fecha"):
            spec[k] = actual.get(k, spec.get(k))
        nombres = sorted(p.name for p in carpeta.glob("img_*.jpg"))
        asegurar_fotos(spec, nombres)
        return jsonify(guardar_resultado(spec, carpeta, id_))
    except ErrorBot as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:
        app.logger.exception("Error inesperado")
        return jsonify(ok=False, error=f"Error inesperado: {e}"), 500


@app.get("/salida/<id_>/<path:archivo>")
def descargar(id_, archivo):
    if not re.fullmatch(r"[\w-]+", id_):
        return "No encontrado", 404
    return send_from_directory(SALIDA / id_, archivo)


if __name__ == "__main__":
    print("=" * 60)
    print(" Bot de informes – Enfoque Ingeniería")
    print(f" Clave de Gemini: {estado_clave()}")
    print(f" Modelo         : {MODELO}")
    print(" Abrí en el navegador: http://127.0.0.1:5000")
    print("=" * 60)

    crear_tablas()
    app.run(host="127.0.0.1", port=5000, debug=False)
