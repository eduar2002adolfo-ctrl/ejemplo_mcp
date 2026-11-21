# =============================================================================
#  Universidad Santiago de Cali - Facultad de Ingeniería
#  Inteligencia Artificial - Práctica: Agentes con MCP + Gemini
#  Parte 5 (Reto) - Paso 1: Servidor MCP de Excel (ExcelServer)
#
#  Integrantes:
#    - David Celorio
#    - Adrián Campos
#    - Martín Murillo
#
#  Herramientas expuestas al modelo:
#    crear_excel, crear_tabla_excel, leer_excel, escribir_celda,
#    agregar_filas, listar_hojas, crear_hoja
#
#  Todas las herramientas devuelven un texto: si algo sale mal, el texto
#  empieza por "Error:" y explica la causa, para que el modelo pueda
#  entender el problema y contárselo al usuario.
# =============================================================================

import os
import re

from mcp.server.mcpserver import MCPServer
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import (column_index_from_string,
                                 coordinate_from_string, range_boundaries)
from openpyxl.utils.exceptions import (CellCoordinatesException,
                                       InvalidFileException)
from openpyxl.worksheet.table import Table, TableStyleInfo

mcp = MCPServer("ExcelServer")

MAX_FILAS_LECTURA = 200          # límite de filas que devuelve leer_excel
CARACTERES_PROHIBIDOS_HOJA = set('[]:*?/\\')


# -----------------------------------------------------------------------------
#  Funciones auxiliares (no son herramientas: el modelo no las ve)
# -----------------------------------------------------------------------------

class ErrorExcel(Exception):
    """Error con un mensaje claro para el usuario."""


def _normalizar_ruta(ruta: str) -> str:
    """Agrega la extensión .xlsx si falta y rechaza otros formatos."""
    ruta = ruta.strip()
    if not ruta:
        raise ErrorExcel("no se indicó el nombre o la ruta del archivo.")
    _, extension = os.path.splitext(ruta)
    if extension == "":
        ruta += ".xlsx"
    elif extension.lower() != ".xlsx":
        raise ErrorExcel(
            f"'{ruta}' no es un archivo .xlsx. Este servidor solo trabaja "
            "con archivos de Excel en formato .xlsx."
        )
    return ruta


def _abrir_libro(ruta: str):
    """Abre un libro existente y devuelve (libro, ruta_normalizada)."""
    ruta = _normalizar_ruta(ruta)
    if not os.path.isfile(ruta):
        raise ErrorExcel(
            f"el archivo '{ruta}' no existe (se buscó en "
            f"'{os.path.abspath(ruta)}'). Revisa el nombre o créalo primero."
        )
    try:
        return load_workbook(ruta), ruta
    except (InvalidFileException, OSError, KeyError, ValueError) as e:
        raise ErrorExcel(
            f"no se pudo abrir '{ruta}'; el archivo está dañado o no es un "
            f"Excel válido ({e})."
        )


def _obtener_hoja(libro, hoja: str):
    """Devuelve la hoja pedida; si hoja está vacía, devuelve la hoja activa."""
    if not hoja:
        return libro.active
    if hoja not in libro.sheetnames:
        disponibles = ", ".join(libro.sheetnames)
        raise ErrorExcel(
            f"la hoja '{hoja}' no existe. Hojas disponibles: {disponibles}."
        )
    return libro[hoja]


def _validar_nombre_hoja(nombre: str) -> str:
    nombre = nombre.strip()
    if not nombre:
        raise ErrorExcel("el nombre de la hoja no puede estar vacío.")
    if len(nombre) > 31:
        raise ErrorExcel(
            f"el nombre de hoja '{nombre}' tiene {len(nombre)} caracteres; "
            "Excel permite máximo 31."
        )
    prohibidos = CARACTERES_PROHIBIDOS_HOJA.intersection(nombre)
    if prohibidos:
        raise ErrorExcel(
            f"el nombre de hoja '{nombre}' contiene caracteres no permitidos: "
            f"{' '.join(sorted(prohibidos))}"
        )
    return nombre


def _validar_celda(celda: str) -> str:
    """Valida una referencia tipo 'A1' y la devuelve en mayúsculas."""
    referencia = celda.strip().upper()
    try:
        columna, fila = coordinate_from_string(referencia)
        indice_columna = column_index_from_string(columna)
    except (CellCoordinatesException, ValueError):
        raise ErrorExcel(
            f"la celda '{celda}' no es válida. Usa el formato columna+fila, "
            "por ejemplo 'A1' o 'C7'."
        )
    if indice_columna > 16384 or fila > 1048576:
        raise ErrorExcel(
            f"la celda '{celda}' está fuera de los límites de Excel "
            "(última celda posible: XFD1048576)."
        )
    return referencia


def _convertir_valor(valor):
    """Convierte textos numéricos ('25', '3.5') a número para que Excel
    los trate como números. Conserva como texto valores con ceros a la
    izquierda (ej. '007', teléfonos o códigos)."""
    if not isinstance(valor, str):
        return valor
    texto = valor.strip()
    if re.fullmatch(r"-?(0|[1-9]\d{0,14})", texto):
        return int(texto)
    if re.fullmatch(r"-?(0|[1-9]\d*)\.\d+", texto):
        return float(texto)
    return valor


def _guardar(libro, ruta: str) -> None:
    try:
        libro.save(ruta)
    except PermissionError:
        raise ErrorExcel(
            f"no se pudo guardar '{ruta}' porque está abierto en otro programa "
            "(por ejemplo Excel) o no hay permisos. Ciérralo e inténtalo de nuevo."
        )


def _ajustar_ancho_columnas(hoja) -> None:
    """Ajusta el ancho de cada columna según su contenido más largo."""
    for columna in hoja.columns:
        largo = max((len(str(c.value)) for c in columna if c.value is not None),
                    default=0)
        letra = get_column_letter(columna[0].column)
        hoja.column_dimensions[letra].width = min(max(largo + 2, 10), 60)


def _hoja_vacia(hoja) -> bool:
    return hoja.max_row == 1 and hoja.max_column == 1 and hoja["A1"].value is None


def _error(e: Exception) -> str:
    return f"Error: {e}"


# -----------------------------------------------------------------------------
#  Herramientas MCP
# -----------------------------------------------------------------------------

@mcp.tool()
def crear_excel(nombre: str, hoja: str = "Hoja1") -> str:
    """Crea un archivo de Excel (.xlsx) nuevo y vacío.

    Args:
        nombre: nombre o ruta del archivo, por ejemplo 'datos.xlsx'
            (si no se escribe la extensión, se agrega .xlsx).
        hoja: nombre de la primera hoja del libro (por defecto 'Hoja1').
    """
    try:
        ruta = _normalizar_ruta(nombre)
        if os.path.exists(ruta):
            raise ErrorExcel(
                f"el archivo '{ruta}' ya existe. Usa otro nombre o trabaja "
                "sobre él con agregar_filas o escribir_celda."
            )
        carpeta = os.path.dirname(ruta)
        if carpeta:
            os.makedirs(carpeta, exist_ok=True)

        libro = Workbook()
        libro.active.title = _validar_nombre_hoja(hoja)
        _guardar(libro, ruta)
        return (f"Archivo de Excel creado en: {os.path.abspath(ruta)} "
                f"(hoja '{libro.active.title}').")
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def crear_tabla_excel(nombre: str, encabezados: list[str],
                      filas: list[list[str]], hoja: str = "Datos") -> str:
    """Crea un archivo de Excel nuevo con una tabla organizada: fila de
    encabezados con formato, los datos debajo y columnas ajustadas.
    Úsala cuando quieras pasar información a una tabla en un Excel nuevo.

    Args:
        nombre: nombre o ruta del archivo nuevo, por ejemplo 'personas.xlsx'.
        encabezados: títulos de las columnas, por ejemplo
            ['Nombre', 'Edad', 'Ciudad'].
        filas: lista de filas; cada fila es una lista de valores en el mismo
            orden que los encabezados, por ejemplo [['Ana', '25', 'Cali']].
        hoja: nombre de la hoja donde se crea la tabla (por defecto 'Datos').
    """
    try:
        ruta = _normalizar_ruta(nombre)
        if os.path.exists(ruta):
            raise ErrorExcel(
                f"el archivo '{ruta}' ya existe. Usa otro nombre para la "
                "tabla nueva o usa agregar_filas para añadir datos a ese archivo."
            )
        encabezados = [str(e).strip() for e in encabezados]
        if not encabezados or any(e == "" for e in encabezados):
            raise ErrorExcel("todos los encabezados deben tener un nombre.")
        for numero, fila in enumerate(filas, start=1):
            if len(fila) > len(encabezados):
                raise ErrorExcel(
                    f"la fila {numero} tiene {len(fila)} valores pero solo hay "
                    f"{len(encabezados)} encabezados."
                )

        libro = Workbook()
        ws = libro.active
        ws.title = _validar_nombre_hoja(hoja)

        ws.append(encabezados)
        for fila in filas:
            completa = list(fila) + [""] * (len(encabezados) - len(fila))
            ws.append([_convertir_valor(v) for v in completa])

        # Formato de encabezados
        for celda in ws[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill("solid", start_color="1F4E78")
            celda.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        _ajustar_ancho_columnas(ws)

        # Tabla de Excel con filtros (requiere encabezados únicos y datos)
        if filas and len(set(encabezados)) == len(encabezados):
            rango = f"A1:{get_column_letter(len(encabezados))}{len(filas) + 1}"
            tabla = Table(displayName="Tabla1", ref=rango)
            tabla.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(tabla)

        carpeta = os.path.dirname(ruta)
        if carpeta:
            os.makedirs(carpeta, exist_ok=True)
        _guardar(libro, ruta)
        return (f"Tabla creada en: {os.path.abspath(ruta)} (hoja '{ws.title}', "
                f"{len(encabezados)} columnas: {', '.join(encabezados)}; "
                f"{len(filas)} filas de datos).")
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def leer_excel(ruta: str, hoja: str = "") -> str:
    """Lee y devuelve el contenido de una hoja de un archivo de Excel,
    fila por fila, con los valores separados por ' | '.

    Args:
        ruta: nombre o ruta del archivo .xlsx.
        hoja: nombre de la hoja a leer; si se deja vacío se lee la hoja activa.
    """
    try:
        libro, ruta = _abrir_libro(ruta)
        ws = _obtener_hoja(libro, hoja)
        if _hoja_vacia(ws):
            return f"La hoja '{ws.title}' de '{ruta}' está vacía."

        lineas = [f"Archivo: {os.path.abspath(ruta)} | Hoja: '{ws.title}' | "
                  f"{ws.max_row} filas x {ws.max_column} columnas "
                  f"(rango {ws.dimensions})"]
        for numero, fila in enumerate(ws.iter_rows(values_only=True), start=1):
            if numero > MAX_FILAS_LECTURA:
                lineas.append(f"... (se muestran solo las primeras "
                              f"{MAX_FILAS_LECTURA} filas)")
                break
            valores = ["" if v is None else str(v) for v in fila]
            lineas.append(f"Fila {numero}: " + " | ".join(valores))
        return "\n".join(lineas)
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def escribir_celda(ruta: str, celda: str, valor: str, hoja: str = "") -> str:
    """Escribe un valor en una celda específica de un archivo de Excel
    existente (reemplaza lo que hubiera en esa celda).

    Args:
        ruta: nombre o ruta del archivo .xlsx.
        celda: referencia de la celda, por ejemplo 'B3'.
        valor: valor a escribir; los números escritos como texto ('25', '3.5')
            se guardan como número.
        hoja: nombre de la hoja; si se deja vacío se usa la hoja activa.
    """
    try:
        referencia = _validar_celda(celda)
        libro, ruta = _abrir_libro(ruta)
        ws = _obtener_hoja(libro, hoja)
        anterior = ws[referencia].value
        ws[referencia] = _convertir_valor(valor)
        _guardar(libro, ruta)
        detalle = f" (antes tenía: {anterior})" if anterior is not None else ""
        return (f"Celda {referencia} de la hoja '{ws.title}' en '{ruta}' "
                f"actualizada con: {valor}{detalle}.")
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def agregar_filas(ruta: str, filas: list[list[str]], hoja: str = "") -> str:
    """Agrega una o varias filas al final de una hoja de un archivo de Excel
    existente. Si la hoja está vacía, la primera fila puede ser de encabezados.

    Args:
        ruta: nombre o ruta del archivo .xlsx.
        filas: lista de filas; cada fila es una lista de valores,
            por ejemplo [['Ana', '25', 'Cali'], ['Luis', '30', 'Bogotá']].
        hoja: nombre de la hoja; si se deja vacío se usa la hoja activa.
    """
    try:
        if not filas:
            raise ErrorExcel("no se recibieron filas para agregar.")
        libro, ruta = _abrir_libro(ruta)
        ws = _obtener_hoja(libro, hoja)

        # Se escribe celda por celda desde la primera fila libre (en una hoja
        # vacía, desde la fila 1).
        ultima_fila_antes = 0 if _hoja_vacia(ws) else ws.max_row
        for i, fila in enumerate(filas, start=1):
            for j, valor in enumerate(fila, start=1):
                ws.cell(row=ultima_fila_antes + i, column=j,
                        value=_convertir_valor(valor))
        ultima_fila = ultima_fila_antes + len(filas)

        # Si la hoja tiene una tabla de Excel que terminaba en la última fila,
        # se extiende para incluir las filas nuevas.
        for tabla in ws.tables.values():
            min_col, min_fila, max_col, max_fila = range_boundaries(tabla.ref)
            if max_fila == ultima_fila_antes:
                tabla.ref = (f"{get_column_letter(min_col)}{min_fila}:"
                             f"{get_column_letter(max_col)}{ultima_fila}")
                if tabla.autoFilter is not None:
                    tabla.autoFilter.ref = tabla.ref

        _ajustar_ancho_columnas(ws)
        _guardar(libro, ruta)
        return (f"Se agregaron {len(filas)} fila(s) a la hoja '{ws.title}' de "
                f"'{ruta}' (filas {ultima_fila_antes + 1} a {ultima_fila}).")
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def listar_hojas(ruta: str) -> str:
    """Lista los nombres de las hojas de un archivo de Excel y cuántas filas
    con datos tiene cada una.

    Args:
        ruta: nombre o ruta del archivo .xlsx.
    """
    try:
        libro, ruta = _abrir_libro(ruta)
        detalle = []
        for ws in libro.worksheets:
            filas = 0 if _hoja_vacia(ws) else ws.max_row
            detalle.append(f"- {ws.title} ({filas} filas)")
        return (f"El archivo '{ruta}' tiene {len(libro.sheetnames)} hoja(s):\n"
                + "\n".join(detalle))
    except ErrorExcel as e:
        return _error(e)


@mcp.tool()
def crear_hoja(ruta: str, nombre_hoja: str) -> str:
    """Agrega una hoja nueva (vacía) a un archivo de Excel existente.

    Args:
        ruta: nombre o ruta del archivo .xlsx.
        nombre_hoja: nombre de la hoja nueva (máximo 31 caracteres,
            sin los caracteres [ ] : * ? / \\).
    """
    try:
        nombre_hoja = _validar_nombre_hoja(nombre_hoja)
        libro, ruta = _abrir_libro(ruta)
        if nombre_hoja in libro.sheetnames:
            raise ErrorExcel(f"ya existe una hoja llamada '{nombre_hoja}'.")
        libro.create_sheet(nombre_hoja)
        _guardar(libro, ruta)
        return (f"Hoja '{nombre_hoja}' creada en '{ruta}'. Hojas actuales: "
                f"{', '.join(libro.sheetnames)}.")
    except ErrorExcel as e:
        return _error(e)


if __name__ == "__main__":
    mcp.run()
