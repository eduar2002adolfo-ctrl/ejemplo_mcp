# =============================================================================
#  Universidad Santiago de Cali - Facultad de Ingeniería
#  Inteligencia Artificial - Práctica: Agentes con MCP + Gemini
#  Parte 2: Servidor MCP de archivos (FileServer)
#
#  Integrantes:
#    - David Celorio
#    - Adrián Campos
#    - Martín Murillo
# =============================================================================

from mcp.server.mcpserver import MCPServer
import os

mcp = MCPServer("FileServer")


@mcp.tool()
def listar_archivos(ruta: str = ".") -> list[str]:
    """ Lista archivos de una carpeta """

    return os.listdir(ruta)


@mcp.tool()
def leer_archivo(ruta: str) -> str:
    """ Lee un archivo de texto """

    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()


@mcp.tool()
def crear_archivo_txt(nombre: str, contenido: str = "") -> str:
    """ Crea un archivo .txt con contenido inicial """

    if not nombre.endswith(".txt"):
        nombre += ".txt"

    with open(nombre, "w", encoding="utf-8") as f:
        f.write(contenido)

    ruta_completa = os.path.abspath(nombre)
    return f"Archivo creado correctamente en: {ruta_completa}"


@mcp.tool()
def agregar_contenido(ruta: str, contenido: str) -> str:
    """ Agrega contenido al final de un archivo txt """

    with open(ruta, "a", encoding="utf-8") as f:
        f.write("\n" + contenido)
    return "Contenido agregado correctamente"


@mcp.tool()
def directorio_actual() -> str:
    """ Retorna el directorio actual del servidor """
    return os.getcwd()


@mcp.tool()
def obtener_ruta(ruta: str) -> str:
    """ Retorna la ruta absoluta de un archivo o carpeta """

    return os.path.abspath(ruta)


@mcp.tool()
def cambiar_directorio(ruta: str) -> str:
    """ Cambia el directorio actual """

    os.chdir(ruta)

    return f"Directorio cambiado a: {os.getcwd()}"


if __name__ == "__main__":
    mcp.run()
