"""
Cliente mínimo de Gemini por REST (solo librería estándar de Python).

Por qué existe: Google ahora emite claves nuevas que empiezan con "AQ." y algunas
versiones del SDK las rechazan con "API key not valid". Llamando directo a la API
con el encabezado x-goog-api-key (forma documentada por Google) funcionan igual.
"""
import json
import os
import time
import urllib.error
import urllib.request

BASE_URL = os.getenv("GEMINI_BASE_URL", "https://generativelanguage.googleapis.com/v1beta")


class ErrorGemini(Exception):
    def __init__(self, codigo, mensaje, clave_invalida=False):
        super().__init__(mensaje)
        self.codigo = codigo
        self.mensaje = mensaje
        self.clave_invalida = clave_invalida


def _pedir(url, clave, cuerpo=None, timeout=180):
    datos = json.dumps(cuerpo).encode("utf-8") if cuerpo is not None else None
    req = urllib.request.Request(
        url, data=datos, method="POST" if cuerpo is not None else "GET",
        headers={"Content-Type": "application/json", "x-goog-api-key": clave},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        texto = e.read().decode("utf-8", errors="replace")
        try:
            err = json.loads(texto).get("error", {})
        except json.JSONDecodeError:
            err = {}
        mensaje = err.get("message") or texto[:300] or str(e)
        invalida = (
            "API key not valid" in texto or "API_KEY_INVALID" in texto
            or "ACCESS_TOKEN_TYPE_UNSUPPORTED" in texto or "UNAUTHENTICATED" in texto
        )
        raise ErrorGemini(e.code, mensaje, invalida)
    except urllib.error.URLError as e:
        raise ErrorGemini(0, f"No se pudo conectar con Google ({e.reason}). Revisá tu conexión a internet.")


def listar_modelos(clave):
    """Devuelve los nombres de modelos que sirven para generar contenido."""
    datos = _pedir(f"{BASE_URL}/models?pageSize=200", clave)
    return [
        m["name"].replace("models/", "")
        for m in datos.get("models", [])
        if "generateContent" in m.get("supportedGenerationMethods", [])
    ]


def generar_texto(clave, modelo, partes, sistema=None, json_salida=True, temperatura=0.3, max_tokens=32768):
    """
    partes: lista de {"text": "..."} o {"inline_data": {"mime_type": "...", "data": "<base64>"}}
    Devuelve el texto generado.
    """
    cuerpo = {
        "contents": [{"role": "user", "parts": partes}],
        "generationConfig": {"temperature": temperatura, "maxOutputTokens": max_tokens},
    }
    if json_salida:
        cuerpo["generationConfig"]["responseMimeType"] = "application/json"
    if sistema:
        cuerpo["system_instruction"] = {"parts": [{"text": sistema}]}

    espera = [4, 12]
    for intento in range(3):
        try:
            datos = _pedir(f"{BASE_URL}/models/{modelo}:generateContent", clave, cuerpo)
            break
        except ErrorGemini as e:
            if e.codigo in (429, 500, 503) and intento < 2:
                time.sleep(espera[intento])
                continue
            raise

    cands = datos.get("candidates") or []
    if not cands:
        motivo = (datos.get("promptFeedback") or {}).get("blockReason", "sin candidatos")
        raise ErrorGemini(200, f"Gemini no devolvió respuesta ({motivo}).")
    cand = cands[0]
    texto = "".join(p.get("text", "") for p in (cand.get("content") or {}).get("parts", []) if not p.get("thought"))
    if cand.get("finishReason") == "MAX_TOKENS":
        raise ErrorGemini(200, "La respuesta salió demasiado larga y se cortó. Probá con menos fotos/datos o pedí un informe más breve.")
    return texto
