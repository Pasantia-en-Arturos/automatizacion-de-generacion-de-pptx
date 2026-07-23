import os
import re
from pptx import Presentation

def reemplazar_texto_manteniendo_formato(parrafo, texto_nuevo: str):
    """Mantiene intacto el estilo original (fuente, color, tamaño y negrita)."""
    if not parrafo.runs:
        parrafo.text = texto_nuevo
        return

    parrafo.runs[0].text = texto_nuevo
    for run in parrafo.runs[1:]:
        run.text = ""


def procesar_parrafo(parrafo, datos_json: dict):
    texto = parrafo.text
    if "{{" not in texto:
        return

    for llave, valor in datos_json.items():
        marcador = f"{{{{{llave}}}}}"
        if marcador in texto:
            texto = texto.replace(marcador, str(valor))

    texto = re.sub(r"\{\{.*?\}\}", "", texto)
    reemplazar_texto_manteniendo_formato(parrafo, texto.strip())


def eliminar_forma(shape):
    """Elimina físicamente un elemento (caja, imagen, círculo) del XML de la diapositiva."""
    sp = shape._element
    sp.getparent().remove(sp)


def eliminar_diapositiva(prs: Presentation, index: int):
    """Elimina físicamente una diapositiva completa."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def eliminar_bloque_visual_del_paso(slide, caja_de_texto_del_paso):
    """
    Elimina TODO el bloque visual asociado a un paso omitido:
    La caja de texto, el círculo con el número del paso y el recuadro gris de imagen.
    """
    # Guardamos la posición vertical (altura en el eje Y) de la caja de texto que vamos a borrar.
    # Usamos un margen de tolerancia (ej. 1 pulgada / ~900,000 unidades EMUs) para atrapar 
    # elementos alineados en esa misma fila aunque estén ligeramente más arriba o abajo.
    posicion_y = caja_de_texto_del_paso.top
    tolerancia_y = 914400 # 1 pulgada en unidades de PowerPoint

    formas_del_bloque = []

    for shape in slide.shapes:
        # Si la forma está a una altura similar en la pantalla, pertenece a este mismo paso
        if abs(shape.top - posicion_y) < tolerancia_y:
            formas_del_bloque.append(shape)

    # Borramos todas las formas que pertenecían a la fila visual de ese paso
    for forma in formas_del_bloque:
        try:
            eliminar_forma(forma)
        except AttributeError:
            pass # Si ya se eliminó previamente en un ciclo anterior, lo ignora


def generar_guia_operativa(ruta_plantilla: str, ruta_salida: str, datos_claude: dict):
    if not datos_claude or not isinstance(datos_claude, dict):
        print("⚠️ ADVERTENCIA: No se recibieron datos de Claude. Operación cancelada.")
        return

    prs = Presentation(ruta_plantilla)
    diapositivas_a_eliminar = []

    for idx, slide in enumerate(prs.slides):
        texto_inicial_slide = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    texto_inicial_slide += p.text

        pasos_en_diapositiva = re.findall(r"\{\{(descripcion_paso_\d+)\}\}", texto_inicial_slide)

        # REGLA 1: Si NINGÚN paso de esta diapositiva vino de Claude, se borra la diapositiva entera
        if pasos_en_diapositiva:
            pasos_activos = [
                paso for paso in pasos_en_diapositiva 
                if paso in datos_claude and str(datos_claude[paso]).strip() != ""
            ]
            if not pasos_activos:
                diapositivas_a_eliminar.append(idx)
                continue

        # REGLA 2: Si la diapositiva se conserva, revisamos si tiene pasos individuales omitidos
        cajas_pasos_a_borrar = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            texto_forma = "".join(p.text for p in shape.text_frame.paragraphs)
            
            # Verificamos si esta caja de texto es de un paso omitido (ej. Paso 6 vacío)
            es_paso_invalido = False
            for paso in re.findall(r"\{\{(descripcion_paso_\d+)\}\}", texto_forma):
                if paso not in datos_claude or str(datos_claude[paso]).strip() == "":
                    es_paso_invalido = True

            if es_paso_invalido:
                cajas_pasos_a_borrar.append(shape)
            else:
                # Si el paso es válido, reemplazamos sus textos y mantenemos su formato
                for paragraph in shape.text_frame.paragraphs:
                    procesar_parrafo(paragraph, datos_claude)

        # Para cada paso que omitió Claude, borramos todo su bloque (texto + círculo + imagen)
        for caja_paso in cajas_pasos_a_borrar:
            eliminar_bloque_visual_del_paso(slide, caja_paso)

    # Eliminamos las diapositivas marcadas como vacías (ej. la página del Paso 7)
    for index in reversed(diapositivas_a_eliminar):
        eliminar_diapositiva(prs, index)

    prs.save(ruta_salida)
    print(f"✅ ¡Éxito! Guía operativa generada en: {ruta_salida}")


# --- PRUEBA CON EL PASO 5 ACTIVO Y PASO 6 VACÍO ---
if __name__ == "__main__":
    directorio_script = os.path.dirname(os.path.abspath(__file__))
    ruta_plantilla = os.path.join(directorio_script, "plantilla.pptx")
    ruta_salida = os.path.join(directorio_script, "Guia_Lista.pptx")

    if os.path.exists(ruta_plantilla):
        datos_prueba = {
            "nombre_proceso": "Conciliación Bancaria Mensual",
            "fecha_vigencia": "Julio 2026",
            "codigo": "FIN-001",
            "revision": "Rev. 2",
            "objetivo": "Garantizar la exactitud del saldo en cuentas bancarias.",
            "norma_1": "Corte bancario al día 30/31 del mes.",
            "responsable_1": "Analista de Tesorería",
            "material_1": "Extracto bancario Excel.",
            "tiempo_ejecucion": "2 horas",
            "contexto_operativo": "Cierre mensual.",
            
            # Pasos activos del 1 al 5
            "descripcion_paso_1": "Descargar estado de cuenta.",
            "descripcion_paso_2": "Importar en el módulo contable.",
            "descripcion_paso_3": "Identificar partidas conciliadas.",
            "descripcion_paso_4": "Elaborar reporte de diferencias.",
            "descripcion_paso_5": "Enviar reporte al supervisor para aprobación final."
            # Al no incluir el paso 6, el script destruirá su caja de texto, su círculo "6" y su recuadro gris.
        }

        generar_guia_operativa(ruta_plantilla, ruta_salida, datos_prueba)