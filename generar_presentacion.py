import json
import re
from pptx import Presentation

def procesar_parrafo(parrafo, datos_json: dict):
    """
    Soluciona el problema 1 y 3:
    Concatena el texto completo del párrafo (fusionando runs)
    y reemplaza únicamente las llaves presentes en datos_json.
    """
    texto_parrafo = parrafo.text
    
    # Si el párrafo no contiene marcadores, no perdemos tiempo procesándolo
    if "{{" not in texto_parrafo:
        return

    # Reemplazamos las llaves que SÍ existan en el diccionario de datos
    for llave, valor in datos_json.items():
        marcador = f"{{{{{llave}}}}}"
        if marcador in texto_parrafo:
            texto_parrafo = texto_parrafo.replace(marcador, str(valor))
    
    # Al asignar p.text, python-pptx colapsa los runs fragmentados en uno solo,
    # conservando la tipografía y estilo a nivel de párrafo.
    parrafo.text = texto_parrafo


def eliminar_diapositiva(prs: Presentation, index: int):
    """Elimina físicamente una diapositiva del archivo PPTX."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def generar_pptx_desde_json(ruta_plantilla: str, ruta_salida: str, datos_json: dict):
    prs = Presentation(ruta_plantilla)
    diapositivas_a_eliminar = []

    # Regex para detectar si quedó algún marcador {{...}} sin reemplazar
    patron_marcador = re.compile(r"\{\{.*?\}\}")

    for i, slide in enumerate(prs.slides):
        marcador_huérfano_encontrado = False

        # Recorremos todas las formas con texto en la diapositiva
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for paragraph in shape.text_frame.paragraphs:
                # 1. Fusionar runs y reemplazar datos de Claude
                procesar_parrafo(paragraph, datos_json)

                # 2. Detectar si quedó algún marcador no usado (ej. {{titulo_3}})
                if patron_marcador.search(paragraph.text):
                    marcador_huérfano_encontrado = True

        # Si la diapositiva conserva marcadores sin llenar, se marca para eliminar (Punto 2)
        if marcador_huérfano_encontrado:
            diapositivas_a_eliminar.append(i)

    # Eliminar diapositivas de atrás hacia adelante para no alterar los índices
    for index in reversed(diapositivas_a_eliminar):
        eliminar_diapositiva(prs, index)

    prs.save(ruta_salida)
    print(f"✅ Presentación generada exitosamente en: {ruta_salida}")


# --- EJEMPLO DE USO ---
if __name__ == "__main__":
    datos_claude = {
        "nombre_proceso": "Conciliación bancaria mensual",
        "titulo_1": "Descarga del estado de cuenta",
        "desc_1": "Ingresar al portal del banco y descargar el estado de cuenta del mes en formato Excel.",
        "titulo_2": "Cruce contra el libro contable",
        "desc_2": "Comparar cada movimiento del estado de cuenta contra los registros del sistema Odoo, identificando diferencias."
    }

    generar_pptx_desde_json("plantilla.pptx", "resultado_final.pptx", datos_claude)