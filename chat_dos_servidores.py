# =============================================================================
#  Universidad Santiago de Cali - Facultad de Ingeniería
#  Inteligencia Artificial - Práctica: Agentes con MCP + Gemini
#  Parte 5 (Reto) - Paso 2: Chat Gemini con DOS servidores MCP
#      - mi_servidor_mcp.py     -> herramientas de archivos y carpetas
#      - servidor_excel_mcp.py  -> herramientas de Excel
#
#  Integrantes:
#    - David Celorio
#    - Adrián Campos
#    - Martín Murillo
# =============================================================================

import asyncio
import os
import sys
from google import genai
from google.genai import types
from mcp import ClientSession
from mcp.client.stdio import (stdio_client, StdioServerParameters)

# Evita que model_copy falle al intentar clonar las tareas asíncronas de la sesión MCP
ClientSession.__deepcopy__ = lambda self, memo: self

# La API Key se toma de la variable de entorno GEMINI_API_KEY.
# (También puede reemplazar el texto "Aquí su llave" por su API Key de Google AI Studio.)
API_KEY = os.environ.get("GEMINI_API_KEY", "Aquí su llave")

MODELO = "gemini-3.1-flash-lite"

# Rutas absolutas: el chat funciona sin importar desde qué carpeta se ejecute.
CARPETA_PROYECTO = os.path.dirname(os.path.abspath(__file__))
# Carpeta de trabajo compartida por los dos servidores: aquí quedan los
# .txt y .xlsx que cree el agente.
CARPETA_TRABAJO = os.path.join(CARPETA_PROYECTO, "archivos_prueba")

# Máximo de llamadas al modelo por cada mensaje del usuario. La prueba final
# (leer varios .txt y luego crear el Excel) necesita más que el valor por
# defecto de la SDK (10).
MAX_LLAMADAS_HERRAMIENTAS = 25

INSTRUCCIONES_SISTEMA = f"""
Eres un asistente técnico experto que trabaja con archivos del sistema y con
hojas de cálculo de Excel. Tienes herramientas de DOS servidores MCP:

1. Servidor de archivos (FileServer): listar_archivos, leer_archivo,
   crear_archivo_txt, agregar_contenido, directorio_actual, obtener_ruta y
   cambiar_directorio. Úsalas para archivos de texto (.txt) y carpetas.
2. Servidor de Excel (ExcelServer): crear_excel, crear_tabla_excel, leer_excel,
   escribir_celda, agregar_filas, listar_hojas y crear_hoja. Úsalas para
   archivos .xlsx.

Reglas:
- Decide tú qué herramientas usar; el usuario no te dirá sus nombres.
- Puedes combinar herramientas de ambos servidores en una misma tarea; por
  ejemplo, leer archivos de texto y luego pasar su información a un Excel.
- Si una tarea necesita varios pasos, hazlos todos antes de responder.
- Para pasar información a una tabla de Excel nueva, usa crear_tabla_excel con
  encabezados claros (uno por cada dato) y una fila por cada archivo o registro.
- Recuerda los archivos que creas durante la conversación para que el usuario
  pueda referirse a ellos sin repetir el nombre ("ese archivo", "los que creaste").
- Los dos servidores empiezan en la misma carpeta de trabajo:
  {CARPETA_TRABAJO}
  Son procesos independientes: si cambias de directorio con el servidor de
  archivos, usa rutas absolutas al trabajar con Excel.
- Si una herramienta responde con un texto que empieza por "Error:", explícale
  el problema al usuario en lenguaje sencillo y, si puedes, corrígelo.
- Responde siempre en español, de forma breve y clara.
"""


def herramientas_del_turno(historial_nuevo) -> list[str]:
    """Devuelve los nombres de las herramientas que el modelo llamó en este turno."""
    nombres = []
    for contenido in historial_nuevo:
        for parte in contenido.parts or []:
            if parte.function_call:
                nombres.append(parte.function_call.name)
    return nombres


# FUNCIÓN PRINCIPAL
async def generate():

    os.makedirs(CARPETA_TRABAJO, exist_ok=True)

    client = genai.Client(api_key=API_KEY)

    # Un StdioServerParameters por cada servidor MCP
    params_archivos = StdioServerParameters(
        command=sys.executable,
        args=[os.path.join(CARPETA_PROYECTO, "mi_servidor_mcp.py")],
        cwd=CARPETA_TRABAJO,
    )
    params_excel = StdioServerParameters(
        command=sys.executable,
        args=[os.path.join(CARPETA_PROYECTO, "servidor_excel_mcp.py")],
        cwd=CARPETA_TRABAJO,
    )

    # CONEXIÓN MCP: se abren los dos servidores y una sesión por cada uno
    async with stdio_client(params_archivos) as (read_archivos, write_archivos), \
            stdio_client(params_excel) as (read_excel, write_excel):

        async with ClientSession(read_archivos, write_archivos) as session_archivos, \
                ClientSession(read_excel, write_excel) as session_excel:

            await session_archivos.initialize()  # Inicializar MCP (archivos)
            await session_excel.initialize()     # Inicializar MCP (Excel)

            # Se muestran las herramientas disponibles de cada servidor
            herramientas_archivos = await session_archivos.list_tools()
            herramientas_excel = await session_excel.list_tools()

            # configuración del chat: las DOS sesiones se pasan juntas a Gemini
            configuracion = types.GenerateContentConfig(
                system_instruction=INSTRUCCIONES_SISTEMA,
                tools=[session_archivos, session_excel],
                thinking_config=types.ThinkingConfig(
                    thinking_level="low",
                    include_thoughts=False
                ),
                temperature=0.8,
                automatic_function_calling=types.AutomaticFunctionCallingConfig(
                    maximum_remote_calls=MAX_LLAMADAS_HERRAMIENTAS
                ),
            )
            chat = client.aio.chats.create(model=MODELO, config=configuracion)

            print("==================================================")
            print("     CHAT MCP + GEMINI  (2 servidores MCP)")
            print("==================================================")
            print("Servidor de archivos:",
                  ", ".join(t.name for t in herramientas_archivos.tools))
            print("Servidor de Excel:   ",
                  ", ".join(t.name for t in herramientas_excel.tools))
            print("Carpeta de trabajo:  ", CARPETA_TRABAJO)
            print("Escribe 'salir' para terminar")

            while True:

                pregunta = input("\n Tú: ")
                if pregunta.lower() in ["salir", "exit"]:
                    break
                if not pregunta.strip():
                    continue
                print("\n Pensando...\n")

                try:
                    mensajes_antes = len(chat.get_history())
                    response = await chat.send_message(pregunta)

                    usadas = herramientas_del_turno(
                        chat.get_history()[mensajes_antes:])
                    if usadas:
                        print(" [Herramientas usadas: " + ", ".join(usadas) + "]\n")

                    print(" Asistente:", response.text)

                except Exception as e:

                    print("\n Error:")
                    print(str(e))


if __name__ == "__main__":
    asyncio.run(generate())
