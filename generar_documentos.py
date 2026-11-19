"""Generador dinámico de archivos de prueba PDF y Excel (.xlsx) a partir de un JSON."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from fpdf import FPDF
from fpdf.enums import XPos, YPos
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

LOREM_IPSUM_PARRAFOS = [
    (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod "
        "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, "
        "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
    ),
    (
        "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore "
        "eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt "
        "in culpa qui officia deserunt mollit anim id est laborum."
    ),
    (
        "Curabitur pretium tincidunt lacus. Nulla gravida orci a odio. Nullam varius, "
        "turpis et commodo pharetra, est eros bibendum elit, nec luctus magna felis "
        "sollicitudin mauris. Integer in mauris eu nibh euismod gravida."
    ),
]

TIPOS_SOPORTADOS = {
    "pdf": "pdf",
    "xlsx": "xlsx",
    "excel": "xlsx",
}


def construir_nombre_version(nombre_base: str, version: int) -> str:
    """Genera el nombre del archivo según la versión (v1 sin sufijo, v2+ con _vN)."""
    if version <= 1:
        return nombre_base
    return f"{nombre_base}_v{version}"


def sanitizar_nombre_hoja_excel(nombre: str) -> str:
    """Limpia caracteres inválidos y recorta a 31 caracteres para nombres de hoja Excel."""
    limpio = re.sub(r"[\[\]:*?/\\]", "_", nombre).strip()
    return (limpio or "Documento")[:31]


def normalizar_tipos(tipo_config: str | list[str]) -> list[str]:
    """Normaliza el campo 'tipo' del JSON a una lista de extensiones soportadas."""
    tipos_crudos = [tipo_config] if isinstance(tipo_config, str) else list(tipo_config)
    resultado: list[str] = []
    for t in tipos_crudos:
        clave = t.strip().lower().lstrip(".")
        if clave not in TIPOS_SOPORTADOS:
            raise ValueError(
                f"Tipo de archivo no soportado: '{t}'. "
                f"Usa 'pdf', 'xlsx' (o 'excel')."
            )
        ext = TIPOS_SOPORTADOS[clave]
        if ext not in resultado:
            resultado.append(ext)
    return resultado


def configurar_fuente_pdf(pdf: FPDF) -> str:
    """Usa Arial (Unicode) si está disponible en Windows; si no, Helvetica integrada."""
    rutas_fuentes = [
        (Path("C:/Windows/Fonts/arial.ttf"), Path("C:/Windows/Fonts/arialbd.ttf")),
        (
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ),
    ]
    for regular, bold in rutas_fuentes:
        if regular.exists() and bold.exists():
            pdf.add_font("CustomSans", "", str(regular))
            pdf.add_font("CustomSans", "B", str(bold))
            return "CustomSans"
    return "Helvetica"


def generar_pdf(ruta_archivo: Path, titulo: str, categoria: str | None = None) -> None:
    """Crea un PDF de una página con el título del documento y texto Lorem Ipsum."""
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_title(titulo)
    pdf.add_page()

    fuente = configurar_fuente_pdf(pdf)

    # Título principal
    pdf.set_font(fuente, "B", 18)
    pdf.set_text_color(30, 41, 59)
    pdf.multi_cell(
        0,
        10,
        titulo,
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Subtítulo opcional con categoría
    if categoria:
        pdf.set_font(fuente, "", 10)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(
            0,
            6,
            f"Categoría: {categoria}  |  Documento de prueba",
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
        )

    # Línea separadora
    pdf.ln(3)
    y_actual = pdf.get_y()
    pdf.set_draw_color(203, 213, 225)
    pdf.line(pdf.l_margin, y_actual, pdf.w - pdf.r_margin, y_actual)
    pdf.ln(6)

    # Contenido Lorem Ipsum
    pdf.set_font(fuente, "", 11)
    pdf.set_text_color(51, 65, 85)
    for parrafo in LOREM_IPSUM_PARRAFOS:
        pdf.multi_cell(
            0,
            7,
            parrafo,
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
        )
        pdf.ln(3)

    pdf.output(str(ruta_archivo))


def generar_excel(ruta_archivo: Path, titulo: str, categoria: str | None = None) -> None:
    """Crea un libro Excel (.xlsx) con una hoja titulada con el nombre del documento y Lorem Ipsum."""
    wb = Workbook()
    wb.properties.title = titulo

    ws = wb.active
    ws.title = sanitizar_nombre_hoja_excel(titulo)

    # Título en A1
    ws["A1"] = titulo
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="1E293B")
    ws["A1"].fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    ws["A1"].alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 30

    fila_actual = 2
    if categoria:
        ws[f"A{fila_actual}"] = f"Categoría: {categoria} | Documento de prueba"
        ws[f"A{fila_actual}"].font = Font(name="Calibri", size=10, italic=True, color="64748B")
        fila_actual += 1

    # Fila en blanco de separación
    fila_actual += 1

    # Párrafos Lorem Ipsum
    for parrafo in LOREM_IPSUM_PARRAFOS:
        celda = ws[f"A{fila_actual}"]
        celda.value = parrafo
        celda.font = Font(name="Calibri", size=11, color="334155")
        celda.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[fila_actual].height = 48
        fila_actual += 1

    ws.column_dimensions["A"].width = 95
    wb.save(str(ruta_archivo))


def generar_desde_config(
    ruta_config: Path,
    carpeta_override: Path | None = None,
    limpiar: bool = False,
) -> list[Path]:
    """Lee el JSON de configuración y genera todos los documentos y versiones solicitados."""
    with ruta_config.open("r", encoding="utf-8") as f:
        config = json.load(f)

    carpeta_destino = (
        carpeta_override
        if carpeta_override is not None
        else Path(config.get("carpeta_destino", "archivos_generados"))
    )
    if not carpeta_destino.is_absolute():
        carpeta_destino = ruta_config.parent / carpeta_destino

    if limpiar and carpeta_destino.exists():
        shutil.rmtree(carpeta_destino)

    carpeta_destino.mkdir(parents=True, exist_ok=True)

    versiones_default = int(config.get("versiones_por_defecto", 2))
    tipo_default = config.get("tipo_por_defecto", "pdf")
    organizar_por_categoria = bool(config.get("organizar_por_categoria", False))
    documentos = config.get("documentos", [])

    if not documentos:
        print("No se encontraron documentos en la configuración.")
        return []

    archivos_creados: list[Path] = []

    for idx, doc in enumerate(documentos, start=1):
        nombre_base = doc.get("nombre", "").strip()
        if not nombre_base:
            raise ValueError(f"El documento #{idx} no tiene un campo 'nombre' válido.")

        categoria = doc.get("categoria")
        versiones = int(doc.get("versiones", versiones_default))
        if versiones < 1:
            raise ValueError(
                f"El documento '{nombre_base}' debe tener al menos 1 versión (recibido: {versiones})."
            )

        tipos = normalizar_tipos(doc.get("tipo", tipo_default))

        dir_salida = carpeta_destino
        if organizar_por_categoria and categoria:
            dir_salida = carpeta_destino / categoria
            dir_salida.mkdir(parents=True, exist_ok=True)

        for version in range(1, versiones + 1):
            nombre_version = construir_nombre_version(nombre_base, version)
            titulo = doc.get("titulo")
            titulo_final = (
                construir_nombre_version(titulo, version) if titulo else nombre_version
            )

            for ext in tipos:
                ruta_archivo = dir_salida / f"{nombre_version}.{ext}"
                if ext == "pdf":
                    generar_pdf(ruta_archivo, titulo_final, categoria)
                elif ext == "xlsx":
                    generar_excel(ruta_archivo, titulo_final, categoria)
                archivos_creados.append(ruta_archivo)
                try:
                    ruta_mostrada = ruta_archivo.relative_to(ruta_config.parent)
                except ValueError:
                    ruta_mostrada = ruta_archivo
                print(f"[OK] {ruta_mostrada}")

    return archivos_creados


def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera archivos de prueba PDF y Excel dinámicamente desde un JSON."
    )
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=Path(__file__).with_name("documentos.json"),
        help="Ruta al archivo JSON de configuración (por defecto: documentos.json).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Sobrescribe la carpeta destino definida en el JSON.",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Elimina el contenido previo de la carpeta destino antes de generar.",
    )
    return parser.parse_args()


def main() -> int:
    args = parsear_argumentos()
    if not args.config.exists():
        print(f"Error: no se encontró el archivo de configuración '{args.config}'.", file=sys.stderr)
        return 1

    archivos = generar_desde_config(
        ruta_config=args.config.resolve(),
        carpeta_override=args.output,
        limpiar=args.clean,
    )
    print(f"\nListo: se generaron {len(archivos)} archivo(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
