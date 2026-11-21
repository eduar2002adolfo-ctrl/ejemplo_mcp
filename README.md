# Práctica: Agentes con MCP + Gemini (Google AI Studio)

**Universidad Santiago de Cali — Facultad de Ingeniería — Inteligencia Artificial**

**Integrantes:** David Celorio · Adrián Campos · Martín Murillo

Agente conversacional que usa Gemini y el protocolo MCP (Model Context Protocol) para
gestionar archivos de texto y, en el reto, archivos de Excel, combinando las herramientas
de dos servidores MCP en una misma conversación.

## Estructura

```
ejemplo_mcp/
├── archivos_prueba/          # carpeta de trabajo: aquí quedan los .txt y .xlsx del reto
├── env_mcp/                  # entorno virtual (se crea en el paso 1)
├── mi_servidor_mcp.py        # Parte 2: servidor MCP de archivos (FileServer)
├── chat_ia.py                # Parte 3: chat Gemini con un servidor MCP
├── servidor_excel_mcp.py     # Parte 5: servidor MCP de Excel (ExcelServer)
├── chat_dos_servidores.py    # Parte 5: chat Gemini con los dos servidores
└── requirements.txt
```

## 1. Configuración (Windows / PowerShell)

```powershell
cd ejemplo_mcp
python -m venv env_mcp
.\env_mcp\Scripts\activate
pip install --upgrade google-genai mcp openpyxl
code .
```

## 2. API Key de Google AI Studio

Cree la llave en https://aistudio.google.com/apikey y guárdela en la terminal
(así no queda escrita en el código):

```powershell
$env:GEMINI_API_KEY = "su_llave_aqui"
```

Opcionalmente, reemplace el texto `"Aquí su llave"` en `chat_ia.py` y `chat_dos_servidores.py`.

## 3. Ejecución

Parte 3 y 4 (un servidor):

```powershell
python chat_ia.py
```

Parte 5 (dos servidores):

```powershell
python chat_dos_servidores.py
```

El chat muestra las herramientas de cada servidor al iniciar y, después de cada respuesta,
la lista de herramientas que el modelo decidió usar (`[Herramientas usadas: ...]`).
Escriba `salir` para terminar.

## 4. Pruebas (sin decirle al asistente qué herramienta usar)

Parte 4 — `chat_ia.py`:

1. Crea un archivo y escribe algo en él.
2. Agrégale otra línea a ese archivo.
3. Muéstrame qué archivos hay en la carpeta y lee el que acabamos de crear.
4. ¿En qué carpeta estás y cuál es la ruta de ese archivo?

Parte 5 — `chat_dos_servidores.py`:

1. Crea al menos 3 archivos de texto, cada uno con información de una persona o un tema
   distinto (por ejemplo: nombre, edad y ciudad, o un libro con su autor y su año).
   Hazlo con contenido diferente en cada uno.
2. Ahora lee los archivos de texto que acabas de crear y organiza su contenido en una
   tabla dentro de un nuevo archivo de Excel.

## Herramientas del servidor de Excel

| Herramienta | Qué hace |
|---|---|
| `crear_excel(nombre, hoja)` | Crea un .xlsx vacío |
| `crear_tabla_excel(nombre, encabezados, filas, hoja)` | Crea un .xlsx nuevo con una tabla con formato |
| `leer_excel(ruta, hoja)` | Devuelve el contenido de una hoja, fila por fila |
| `escribir_celda(ruta, celda, valor, hoja)` | Escribe un valor en una celda (ej. `B3`) |
| `agregar_filas(ruta, filas, hoja)` | Agrega filas al final de una hoja |
| `listar_hojas(ruta)` | Lista las hojas del libro |
| `crear_hoja(ruta, nombre_hoja)` | Agrega una hoja nueva |

Errores controlados: archivo inexistente, archivo que no es .xlsx, hoja inexistente,
celda inválida, nombre de hoja inválido, archivo ya existente y archivo abierto en Excel.
