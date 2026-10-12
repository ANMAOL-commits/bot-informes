# Bot de informes – Enfoque Ingeniería

Página HTML local donde cargás **datos, imágenes y audios** y el bot (Gemini, gratis) redacta el informe.
Descarga en **Word (.docx)** y **HTML**, con el formato del modelo (encabezado, secciones numeradas, anexo fotográfico, firmas).
Lo que redacta la IA sale **resaltado en amarillo** para que lo revises.

## Paso a paso (Windows + Python 3.14 64 bits + Visual Studio Code)

### 1. Verificar Python
Abrí PowerShell y ejecutá:
    python --version
    python -c "import struct; print(struct.calcsize('P')*8)"
Debe mostrar `Python 3.14.x` y `64`.

### 2. Preparar la carpeta en VS Code
1. Descomprimí este proyecto en `C:\Bot_informes`.
2. En VS Code: **Archivo > Abrir carpeta** → `C:\Bot_informes`.
3. Instalá la extensión **Python** (de Microsoft) si no la tenés.
4. Abrí la terminal integrada: **Terminal > Nueva terminal**.

### 3. Crear el entorno virtual e instalar librerías
    python -m venv .venv
    .venv\Scripts\Activate.ps1
    pip install --upgrade pip
    pip install -r requirements.txt
Si PowerShell dice que la ejecución de scripts está deshabilitada:
    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
y repetí la activación. Luego `Ctrl+Shift+P` > **Python: Select Interpreter** > elegí el que dice `.venv`.

### 4. Conseguir la clave de Gemini
1. Entrá a https://aistudio.google.com/api-keys con tu cuenta de Google.
2. **Crear clave de API** (o usá la existente del proyecto) y copiá la clave completa con el botón de copiar.
3. **Las claves nuevas empiezan con `AQ.`** (desde mayo de 2026 AI Studio solo emite este formato). Es normal y válida; las viejas `AIza...` también sirven si siguen activas.
4. Copiá el archivo `.env.example` como `.env` y pegá la clave:
       GEMINI_API_KEY=AQ.tu_clave_completa
   Sin comillas, sin espacios, sin otras líneas pegadas.

### 5. Probar la clave (importante)
    python probar_clave.py
Debe terminar con `✔ Todo funciona`. Este bot llama a Google directamente (sin la librería `google-genai`),
porque algunas versiones de esa librería rechazan las claves `AQ.` con "API key not valid".
Si el modelo no existe, el script lista los disponibles: copiá uno en `GEMINI_MODEL` dentro de `.env`.

### 6. Ejecutar el bot
    python app.py
Abrí http://127.0.0.1:5000 (o hacé doble clic en `iniciar_bot.bat` después de tener el entorno creado).

### 7. Usarlo
1. Elegí el tipo de informe y completá obra, código y fecha.
2. Escribí datos en el cuadro de texto, subí fotos y/o audios (los de WhatsApp en .opus sirven).
3. **Generar informe** → mirá la vista previa, descargá el Word.
4. Para cambios: escribí la instrucción en "Pedir cambios" (ej. "ampliá riesgos y sacá la foto 3").

El checkbox **Modo prueba** arma el informe sin usar la IA (sirve para verificar la instalación).

## Logo
Guardá tu logo como `static/logo.png`; aparece en el encabezado del Word.

## Límites y cuidados
- Límites: Gemini limita las consultas por minuto y por día según tu nivel. Si aparece el aviso de límite, esperá y reintentá. Si tenés facturación prepaga, cada informe descuenta un costo pequeño de tu crédito (revisalo en AI Studio > Uso / Gasto).
- Según los términos de Google, en el nivel gratuito el contenido enviado puede usarse para mejorar sus productos. Verificá los términos vigentes y no subas material confidencial de clientes sin revisarlo.
- Audios: hasta ~15 MB en total por informe. Fotos: hasta 20.
- La IA redacta; el profesional responsable revisa y firma. Nunca confiar en cantidades o normas generadas sin verificar.
- No compartas ni subas a internet el archivo `.env`.

## Estructura
    app.py              servidor y armado de consultas
    gemini_rest.py      llamada directa a la API de Gemini (sin SDK)
    generador_docx.py   arma el Word
    vista_html.py       arma la versión HTML
    plantillas/         estructura de cada tipo de informe (JSON)
    templates/index.html  la página
    probar_clave.py     diagnóstico de la clave
    salida/             informes generados (una carpeta por informe)

## Errores frecuentes
| Mensaje | Causa y solución |
|---|---|
| API key not valid / API_KEY_INVALID | La clave del `.env` está incompleta, mal copiada o fue borrada: copiala de nuevo desde AI Studio, sin comillas, y reiniciá `app.py` |
| FALTA: no hay clave en el archivo .env | El `.env` no existe o está en otra carpeta; debe estar junto a `app.py` |
| 429 / límite gratuito | Demasiadas consultas: esperá un minuto |
| 404 modelo no existe | Ejecutá `probar_clave.py` y poné un modelo de la lista en `GEMINI_MODEL` |
| Formato de audio no soportado | Convertí a mp3 o wav |
