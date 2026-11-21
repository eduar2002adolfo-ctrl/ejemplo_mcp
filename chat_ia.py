# =============================================================================
#  Universidad Santiago de Cali - Facultad de Ingeniería
#  Inteligencia Artificial - Práctica: Agentes con MCP + Gemini
#  Parte 3: Chat LLM Gemini conectado a un servidor MCP (archivos)
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


# FUNCIÓN PRINCIPAL
async def generate():

    client = genai.Client(api_key=API_KEY)

    server_params = StdioServerParameters(
        command=sys.executable,
        args=["mi_servidor_mcp.py"],
    )

    # CONEXIÓN MCP
    async with stdio_client(server_params) as (read, write):

        async with ClientSession(read, write) as session:
            await session.initialize()  # Inicializar MCP

            # Instrucciones del sistema
            instruciones_sistema = (
                "Eres un asistente técnico experto. Usa las herramientas MCP siempre que "
                "el usuario quiera trabajar con archivos o directorios."
            )

            # configuración del chat, se le indica que use las herramientas MCP
            configuracion = types.GenerateContentConfig(
                system_instruction=instruciones_sistema,
                tools=[session],
                thinking_config=types.ThinkingConfig(
                    thinking_level="low",
                    include_thoughts=False
                ),
                temperature=0.8,
            )
            # modelo a usar
            modelo = "gemini-3.1-flash-lite"
            chat = client.aio.chats.create(model=modelo, config=configuracion)

            print("====================================")
            print("       CHAT MCP + GEMINI")
            print("====================================")
            print("Escribe 'salir' para terminar")

            while True:

                pregunta = input(" Tú: ")
                if pregunta.lower() in ["salir", "exit"]:
                    break
                print("\n Pensando...\n")

                try:
                    response = await chat.send_message(
                        pregunta
                    )
                    print(response.text)

                except Exception as e:

                    print("\n Error:")
                    print(str(e))


if __name__ == "__main__":
    asyncio.run(generate())
