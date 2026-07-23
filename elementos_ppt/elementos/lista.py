"""
Elemento de lista
=================

Combina un punto de viñeta + un cuadro de texto en un solo bloque, que es el
patrón repetido en Normas, Responsables, Materiales, Tips, etc.
"""

from pptx.util import Emu
from .seccion import crear_punto_vineta
from .texto import crear_cuadro_texto


def crear_item_lista(slide, left, top, width, height, marcador):
    """Crea el punto de viñeta y su caja de texto asociada, alineados en la
    misma fila. Devuelve (punto, caja) para poder agruparlos después si se
    quiere (ver ejemplo en demo.py)."""
    diametro = Emu(220200)
    punto = crear_punto_vineta(
        slide,
        left=left - diametro - Emu(90000),
        top=top + (height - diametro) // 2,
        diametro=diametro,
    )
    caja = crear_cuadro_texto(slide, texto="", left=left, top=top,
                               width=width, height=height, marcador=marcador)
    return punto, caja
