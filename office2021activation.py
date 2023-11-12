import xml.etree.ElementTree as ET
import xml.dom.minidom
import glob
import os


PRODUCTS_KEYS = {
    # *    Pos_elem     :           Producto,Clave
    1: "ProPlus2021Volume,FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
    2: [
        "Word2021Volume,TN8H9-M34D3-Y64V9-TR72V-X79KV",
        "Excel2021Volume,NWG3X-87C9K-TC7YY-BC2G7-G6RVC",
        "PowerPoint2021Volume,TY7XF-NFRBR-KJ44C-G83KF-GX27K",
    ],
    3: "Word2021Volume,TN8H9-M34D3-Y64V9-TR72V-X79KV",
    4: "Excel2021Volume,NWG3X-87C9K-TC7YY-BC2G7-G6RVC",
    5: "PowerPoint2021Volume,TY7XF-NFRBR-KJ44C-G83KF-GX27K",
    6: "Access2021Volume,WM8YG-YNGDD-4JHDC-PG3F4-FC4T4",
    7: "Publisher2021Volume,2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ",
    8: "SkypeforBusiness2021Volume,HWCXN-K3WBT-WJBKY-R8BD9-XK29P",
    9: "Outlook2021Volume,C9FM6-3N72F-HFJXB-TM3V9-T86R9",
}


# Buscar el archivo ✅
def find_file():
    # Patrón de búsqueda
    search_pattern = "**/Office2021/configuration.xml"
    # Buscar xml
    xml_file = glob.glob(search_pattern, recursive=True)

    if xml_file:
        os.system("cls")
        return xml_file[0]
    else:
        print("El archivo 'configuration.xml' no se encontró.")
        return None


# Elegir lo que quieres descargar ✅
def opt_download() -> None:
    if find_file():
        print(
            """
            *************************************************
            * INSTALACION Y ACTIVACION LEGAL DE OFFICE 2021 * 
            *                    v 1.4.1                    *
            *************************************************

            ¿Que quieres instalar?
            1) Todo el pack (Suite)
            2) Word, Excel y PowerPoint
            3) Word
            4) PowerPoint
            5) Excel
            6) Access
            7) Publisher
            8) Skype Empresarial
            9) Outlook
        """
        )
        a = int(input("¿Que quieres descargar? > "))
        lang_prod(a)


# Prerequisitos ✅
def lang_prod(a: int) -> None:
    b = int(input("¿Quieres instalar en español (1) o en ingles (2)? > "))

    if a in range(1, 10):
        save_xml(a, b, PRODUCTS_KEYS.get(a))
    else:
        print("Revise su sintaxis")


# Insertar el idioma, el nombre del producto y la clave ✅
def save_xml(a: int, b: int, _: list[str]) -> None:
    if a != 2:
        pid, pidkey, idlang = (
            _.split(",")[0],
            _.split(",")[1],
            "es-es" if b == 1 else "en-us",
        )
        PRODUCTS_KEYS_SELECTED = f'<Product ID="{pid}" PIDKEY="{pidkey}">\n    <Language ID="{idlang}" />\n</Product>'
        insert_xml_element(PRODUCTS_KEYS_SELECTED)
    else:
        for item in _:
            pid, pidkey, idlang = (
                item.split(",")[0],
                item.split(",")[1],
                "es-es" if b == 1 else "en-us",
            )
            PRODUCTS_KEYS_SELECTED = f'<Product ID="{pid}" PIDKEY="{pidkey}">\n    <Language ID="{idlang}" />\n</Product>'
            insert_xml_element(PRODUCTS_KEYS_SELECTED)

    exec_install()


# Guardar en XML ✅
def insert_xml_element(xml_str: str) -> None:
    with open(find_file(), "r+", encoding="utf-8") as file:
        # Lee el contenido actual del archivo
        file_content = file.readlines()
        # Inserta la cadena
        file_content.insert(3, (xml_str + "\n"))
        # Vuelve al inicio del archivo y escribe el contenido actualizado
        file.seek(0)
        file.writelines(file_content)
        file.truncate()


# Ejecutar sentencia ✅
def exec_install() -> None:
    # Patrón de búsqueda
    ruta = "**/Office2021/setup.exe"
    # Buscar exe
    exe_file = glob.glob(ruta, recursive=True)

    if exe_file:
        os.system(f"{exe_file[0]} /configure {find_file()}")
    else:
        print("El archivo 'setup.exe' no se encontró.")


############################################################################
# Inicio del programa
opt_download()
############################################################################
