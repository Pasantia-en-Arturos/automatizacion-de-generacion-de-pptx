"""
Estilo de marca
===============

Valores extraídos directamente de la plantilla "Guía Operativa" de RRHH
(colores del tema y tipografías reales), para que los elementos creados por
código combinen con el diseño oficial en vez de inventar valores.
"""

from pptx.dml.color import RGBColor

NARANJA = RGBColor(0xF2, 0x6F, 0x33)      # títulos y acentos (accent del tema)
GRIS_CUERPO = RGBColor(0x59, 0x59, 0x59)  # texto de cuerpo (dk2 del tema)
GRIS_BORDE = RGBColor(0xDB, 0xDC, 0xDD)   # borde de los recuadros de sección
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

FUENTE_TITULO = "Montserrat ExtraBold"
FUENTE_CUERPO = "Poppins SemiBold"

TAMANO_TITULO_PRINCIPAL = 72
TAMANO_TITULO_SECCION = 48
TAMANO_CUERPO = 31.2
