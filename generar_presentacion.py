import os
import re
from pptx import Presentation

def procesar_parrafo(parrafo, datos_json: dict):
    """
    Fusiona los fragmentos de texto (runs) y reemplaza las llaves de Claude.
    """
    texto = parrafo.text
    if "{{" not in texto:
        return

    # 1. Reemplazamos los valores que SÍ envió Claude
    for llave, valor in datos_json.items():
        marcador = f"{{{{{llave}}}}}"
        if marcador in texto:
            texto = texto.replace(marcador, str(valor))
    
    # 2. Limpieza suave: Si quedó alguna etiqueta sobrante como {{tip_3}} o {{norma_3}},
    # la borramos silenciosamente (la reemplazamos por texto vacío "") para no dañar la forma.
    texto = re.sub(r"\{\{.*?\}\}", "", texto)
    
    # Asignamos el texto limpio, conservando tipografía y formato original
    parrafo.text = texto.strip()


def eliminar_diapositiva(prs: Presentation, index: int):
    """Elimina físicamente una diapositiva del archivo PPTX."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def generar_guia_operativa(ruta_plantilla: str, ruta_salida: str, datos_claude: dict):
    prs = Presentation(ruta_plantilla)
    diapositivas_a_eliminar = []

    for idx, slide in enumerate(prs.slides):
        # Recopilamos todo el texto de la diapositiva ANTES de procesarla
        texto_inicial_slide = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    texto_inicial_slide += p.text

        # Detectamos qué pasos específicos (ej. descripcion_paso_5) viven en esta diapositiva
        pasos_en_diapositiva = re.findall(r"\{\{(descripcion_paso_\d+)\}\}", texto_inicial_slide)

        # LÓGICA DE BORRADO DE DIAPOSITIVAS:
        # Si la diapositiva contiene pasos del procedimiento, verificamos si al menos UNO
        # de esos pasos vino en el JSON de Claude con texto útil.
        if pasos_en_diapositiva:
            paso_activo_encontrado = any(
                paso in datos_claude and str(datos_claude[paso]).strip() != "" 
                for paso in pasos_en_diapositiva
            )
            
            # Si ningún paso de esta diapositiva tiene datos, la marcamos para eliminar
            if not paso_activo_encontrado:
                diapositivas_a_eliminar.append(idx)
                continue # Saltamos al siguiente slide sin procesar este

        # Si la diapositiva es válida, procedemos a reemplazar el contenido de sus formas
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    procesar_parrafo(paragraph, datos_claude)

    # Eliminamos las diapositivas vacías de atrás hacia adelante
    for index in reversed(diapositivas_a_eliminar):
        eliminar_diapositiva(prs, index)

    prs.save(ruta_salida)
    print(f"✅ ¡Éxito! Guía operativa generada en: {ruta_salida}")


# --- PRUEBA CON DATOS REALES DE TU ESTRUCTURA ---
if __name__ == "__main__":
    directorio_script = os.path.dirname(os.path.abspath(__file__))
    ruta_plantilla = os.path.join(directorio_script, "plantilla.pptx")
    ruta_salida = os.path.join(directorio_script, "Guia_Lista.pptx")

    if not os.path.exists(ruta_plantilla):
        print(f"❌ ERROR: No se encontró 'plantilla.pptx' en la carpeta:\n   {directorio_script}")
    else:
        # Diccionario simulando la respuesta exacta de Claude adaptada a tu plantilla
        json_real_claude = {
            # Datos Generales (Diapositiva 1)
            "nombre_proceso": "Conciliación Bancaria Mensual",
            "fecha_vigencia": "Julio 2026",
            "codigo": "FIN-001",
            "revision": "Rev. 2",
            "objetivo": "Garantizar la exactitud del saldo en cuentas bancarias frente al libro mayor.",
            "norma_1": "No iniciar sin el corte bancario al día 30/31 del mes.",
            "norma_2": "Todos los ajustes deben estar aprobados por el supervisor.",
            "norma_3": "", # Si no aplica una tercera norma, se limpia sola
            "responsable_1": "Analista de Tesorería",
            "responsable_2": "Coordinador de Contabilidad",
            "material_1": "Extracto bancario en formato Excel o PDF.",
            "material_2": "Reporte auxiliar de bancos exportado de Odoo.",
            "equipo_1": "Ninguno requerido.",
            "equipo_2": "",
            "tip_1": "Verificar primero las partidas de mayor monto para ganar tiempo.",
            "tip_2": "Utilizar la función BUSCARV para conciliar movimientos masivos.",
            "tip_3": "",
            "tiempo_ejecucion": "2 a 3 horas por cuenta",
            "titulo_conocimiento_1": "Manejo de Odoo",
            "texto_conocimiento_1": "Saber exportar libros auxiliares e identificar asientos.",
            "titulo_conocimiento_2": "Excel Intermedio",
            "texto_conocimiento_2": "Dominio de tablas dinámicas y filtros avanzados.",
            "titulo_conocimiento_3": "",
            "texto_conocimiento_3": "",
            
            # Procedimiento (Diapositiva 2 en adelante)
            "contexto_operativo": "Proceso crítico de cierre mensual en el área de finanzas.",
            "descripcion_paso_1": "Ingresar al portal bancario con credenciales de lectura y descargar el estado de cuenta.",
            "descripcion_paso_2": "Importar el extracto en el módulo contable y ejecutar el cruce automático.",
            "descripcion_paso_3": "Identificar las partidas conciliadas y marcar los movimientos en tránsito.",
            "descripcion_paso_4": "Elaborar el reporte de diferencias y enviarlo para revisión."
            # NOTA: Omitimos intencionalmente los pasos 5, 6 y 7 para probar que 
            # el script elimine automáticamente las últimas diapositivas.
        }

        generar_guia_operativa(ruta_plantilla, ruta_salida, json_real_claude)