import os

from openpyxl import Workbook, load_workbook
from tabulate import tabulate

FILE_NAME = "personas.xlsx"
DEFAULT_HEADERS = ["Id", "Name", "Company", "Email", "MAC Address"]


def create_sample_workbook(file_path: str = FILE_NAME) -> None:
    """Crea un archivo de ejemplo si no existe."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Personas"

    data = [
        [1, "Ana Gómez", "Group One", "ana@groupone.com", "00:1A:2B:3C:4D:5E"],
        [2, "Luis Ruiz", "Alpha Tech", "luis@alphatech.com", "00:22:33:44:55:66"],
        [3, "Marta Pérez", "Tech Group", "marta@techgroup.com", "00:AA:BB:CC:DD:EE"],
        [4, "Pedro Vega", "Group Dynamics", "pedro@groupdynamics.com", "10:20:30:40:50:60"],
        [5, "Sofía Torres", "Prism Labs", "sofia@prismlabs.com", "20:30:40:50:60:70"],
        [6, "Javier López", "Nova Group", "javier@novagroup.com", "30:40:50:60:70:80"],
    ]

    sheet.append(DEFAULT_HEADERS)
    for row in data:
        sheet.append(row)

    workbook.save(file_path)
    print(f"Se creó el archivo {file_path} con datos de ejemplo.")


def load_personas(file_path: str = FILE_NAME):
    """Carga el archivo Excel y devuelve los datos sin encabezados."""
    if not os.path.exists(file_path):
        create_sample_workbook(file_path)

    workbook = load_workbook(file_path)
    sheet = workbook.active

    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and any(cell is not None and str(cell).strip() != "" for cell in row):
            data.append(list(row))

    return data


def print_table(rows, headers=None):
    """Muestra una tabla formateada en consola."""
    if headers is None:
        headers = DEFAULT_HEADERS

    if not rows:
        print("No hay registros para mostrar.")
        return

    table = [headers] + rows
    print(tabulate(table, headers="firstrow", tablefmt="fancy_grid"))


def reto_1_nombre_email(rows):
    """Reto 1: mostrar únicamente Name y Email."""
    print("\nReto 1: Nombre y Email")
    table = [["Name", "Email"]]
    for row in rows:
        table.append([row[1], row[3]])
    print(tabulate(table, headers="firstrow", tablefmt="fancy_grid"))


def reto_2_contar_registros(rows):
    """Reto 2: contar cuántos registros existen."""
    total = len(rows)
    print(f"\nReto 2: Hay {total} registros en el archivo.")


def reto_3_filtrar_empresa(rows, keyword="Group"):
    """Reto 3: filtrar personas cuya empresa contenga la palabra Group."""
    print(f"\nReto 3: Filtrar empresas que contengan '{keyword}'")
    filtrados = [row for row in rows if keyword.lower() in str(row[2]).lower()]

    if not filtrados:
        print("No se encontraron coincidencias.")
        return

    table = [["Id", "Name", "Company", "Email"]]
    for row in filtrados:
        table.append([row[0], row[1], row[2], row[3]])
    print(tabulate(table, headers="firstrow", tablefmt="fancy_grid"))


def reto_4_buscar_por_id(rows):
    """Reto 4: pedir un ID y mostrar una persona."""
    while True:
        try:
            valor = int(input("Ingrese el ID a buscar: "))
            break
        except ValueError:
            print("Error: debe ingresar un número válido.")

    persona = None
    for row in rows:
        if int(row[0]) == valor:
            persona = row
            break

    if persona is None:
        print(f"No se encontró ninguna persona con el ID {valor}.")
        return

    print("\nReto 4: Persona encontrada")
    tabla = [["Id", "Name", "Company", "Email", "MAC Address"]]
    tabla.append(persona)
    print(tabulate(tabla, headers="firstrow", tablefmt="fancy_grid"))


def mostrar_todos(rows):
    """Muestra toda la información en forma tabular."""
    print("\nListado completo de personas")
    print_table(rows, DEFAULT_HEADERS)


def buscar_por_id(rows):
    """Busca una persona por ID con validación."""
    while True:
        try:
            valor = int(input("Ingrese el ID a buscar: "))
            break
        except ValueError:
            print("Error: debe ingresar un número entero.")

    resultado = [row for row in rows if int(row[0]) == valor]

    if not resultado:
        print(f"No existe una persona con el ID {valor}.")
        return

    print(f"\nPersona con ID {valor}")
    print_table(resultado, DEFAULT_HEADERS)


def filtrar_empresa(rows):
    """Filtra personas por nombre de empresa."""
    palabra = input("Ingrese la palabra a buscar en la empresa: ").strip()
    resultado = [row for row in rows if palabra.lower() in str(row[2]).lower()]

    if not resultado:
        print(f"No se encontraron personas cuya empresa contenga '{palabra}'.")
        return

    print(f"\nResultados para la empresa que contiene '{palabra}'")
    print_table(resultado, DEFAULT_HEADERS)


def menu_principal():
    """Menú interactivo principal."""
    personas = load_personas(FILE_NAME)

    while True:
        print("\n====================================")
        print("MENÚ DE CONSULTA DE PERSONAS")
        print("====================================")
        print("1. Ver todos")
        print("2. Buscar por ID")
        print("3. Filtrar empresa")
        print("4. Contar registros")
        print("5. Reto 1: Nombre y Email")
        print("6. Reto 3: Filtrar Group")
        print("7. Reto 4: Buscar por ID")
        print("8. Salir")

        opcion = input("Seleccione una opción: ").strip()

        if opcion == "1":
            mostrar_todos(personas)
        elif opcion == "2":
            buscar_por_id(personas)
        elif opcion == "3":
            filtrar_empresa(personas)
        elif opcion == "4":
            reto_2_contar_registros(personas)
        elif opcion == "5":
            reto_1_nombre_email(personas)
        elif opcion == "6":
            reto_3_filtrar_empresa(personas, "Group")
        elif opcion == "7":
            reto_4_buscar_por_id(personas)
        elif opcion == "8":
            print("Gracias por usar el sistema.")
            break
        else:
            print("Opción no válida. Intente nuevamente.")


if __name__ == "__main__":
    print("=== Taller Python - Lectura de Excel ===")
    menu_principal()
