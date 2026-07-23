import re
from pptx import Presentation

def procesar_parrafo(parrafo, datos_json: dict) -> bool:
    """
    Soluciona el problema 1 y 3:
    Concatena los 'runs' fragmentados y reemplaza los valores de Claude.
    Devuelve True si al final de todo quedó un marcador {{...}} sin reemplazar.
    """
    texto_parrafo = parrafo.text
    
    # Si no hay marcadores en este párrafo, terminamos rápido
    if "{{" not in texto_parrafo:
        return False

    # Reemplazamos las llaves por el contenido del JSON de Claude
    for llave, valor in datos_json.items():
        marcador = f"{{{{{llave}}}}}"
        if marcador in texto_parrafo:
            # Reemplazo estricto tratando todo como texto
            texto_parrafo = texto_parrafo.replace(marcador, str(valor))
    
    # AL SOBRESCRIBIR .text:
    # 1. Se fusionan automáticamente todos los 'runs' internos rotos de PowerPoint.
    # 2. Se hereda exactamente el único formato (fuente, color, tamaño) que tenía asignado.
    parrafo.text = texto_parrafo

    # Verificamos si sobrevivió alguna llave huérfana (ej. {{titulo_3}} que Claude no envió)
    return bool(re.search(r"\{\{.*?\}\}", texto_parrafo))


def eliminar_forma(shape):
    """Elimina una caja de texto o forma específica de la diapositiva."""
    sp = shape._element
    sp.getparent().remove(sp)


def eliminar_diapositiva(prs: Presentation, index: int):
    """Elimina físicamente una diapositiva completa del archivo PPTX."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def generar_presentacion_inteligente(ruta_plantilla: str, ruta_salida: str, datos_claude: dict):
    prs = Presentation(ruta_plantilla)
    diapositivas_a_eliminar = []

    for idx_slide, slide in enumerate(prs.slides):
        formas_a_eliminar = []
        cajas_de_texto_validas = 0

        # NIVEL 1: Analizar forma por forma dentro de la diapositiva
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            forma_tiene_huérfano = False
            texto_total_forma = ""

            for paragraph in shape.text_frame.paragraphs:
                # Procesamos el párrafo
                quedó_huerfano = procesar_parrafo(paragraph, datos_claude)
                if quedó_huerfano:
                    forma_tiene_huérfano = True
                texto_total_forma += paragraph.text.strip()

            # Si la caja de texto se quedó con una llave como {{titulo_3}}, la marcamos para borrar
            if forma_tiene_huérfano:
                formas_a_eliminar.append(shape)
            elif len(texto_total_forma) > 0:
                # Contamos cuántas cajas de texto reales y con contenido útil sobreviven
                cajas_de_texto_validas += 1

        # Borramos físicamente las cajas de texto que sobraron en esta diapositiva
        for shape in formas_a_eliminar:
            eliminar_forma(shape)

        # NIVEL 2: Si al borrar los pasos sobrantes la diapositiva se quedó sin contenido
        # (cajas de texto útiles), la marcamos para eliminar la diapositiva completa.
        if cajas_de_texto_validas == 0:
            diapositivas_a_eliminar.append(idx_slide)

    # Eliminar diapositivas sobrantes DE ATRÁS HACIA ADELANTE para no romper los índices
    for index in reversed(diapositivas_a_eliminar):
        eliminar_diapositiva(prs, index)

    prs.save(ruta_salida)
    print(f"✅ ¡Éxito! Presentación guardada en: {ruta_salida}")


# --- PRUEBA CON EL JSON DE CLAUDE ---
if __name__ == "__main__":
    json_de_claude = {
        "nombre_proceso": "Conciliación bancaria mensual",
        "titulo_1": "Descarga del estado de cuenta",
        "desc_1": "Ingresar al portal del banco y descargar el estado de cuenta del mes en formato Excel.",
        "titulo_2": "Cruce contra el libro contable",
        "desc_2": "Comparar cada movimiento del estado de cuenta contra los registros del sistema Odoo, identificando diferencias."
        # Notar que intencionalmente no enviamos titulo_3 ni desc_3
    }

    generar_presentacion_inteligente("plantilla.pptx", "presentacion_lista.pptx", json_de_claude)