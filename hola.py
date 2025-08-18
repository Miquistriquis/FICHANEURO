import pandas as pd
import json
import os
import traceback
from tkinter import Tk, filedialog
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
# --- CAMBIO 1: Importamos el módulo 'units' completo ---
from reportlab.lib import units
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.colors import navy, black, crimson

def seleccionar_archivo_excel():
    """Abre un explorador para que el usuario seleccione el archivo Excel."""
    root = Tk()
    root.withdraw()
    ruta_archivo = filedialog.askopenfilename(
        title="Selecciona el archivo Excel con los registros",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )
    return ruta_archivo

def seleccionar_directorio_salida():
    """Abre un explorador para que el usuario seleccione la carpeta de guardado."""
    root = Tk()
    root.withdraw()
    ruta_directorio = filedialog.askdirectory(
        title="Selecciona la carpeta donde se guardarán los PDFs"
    )
    return ruta_directorio

def calcular_puntaje_aq10(respuestas_alumno, plantilla_seccion):
    """Calcula el puntaje del tamizaje AQ-10."""
    puntaje = 0
    items_positivos = [1, 7, 8, 10]
    opciones_acuerdo = ["(3) Un poco de acuerdo", "(4) Definitivamente de acuerdo"]
    opciones_desacuerdo = ["(1) Definitivamente, en desacuerdo", "(2) Un poco en desacuerdo"]
    
    for i, pregunta in enumerate(plantilla_seccion['preguntas'], 1):
        id_pregunta = pregunta['id']
        respuesta = respuestas_alumno.get(id_pregunta, "")
        
        if i in items_positivos:
            if any(op in respuesta for op in opciones_acuerdo):
                puntaje += 1
        else:
            if any(op in respuesta for op in opciones_desacuerdo):
                puntaje += 1
            
    interpretacion = "Puntaje dentro del rango esperado."
    if puntaje >= 6:
        interpretacion = "Puntaje sugestivo. Se recomienda una valoración más profunda por un especialista."
        
    return puntaje, interpretacion

def calcular_puntaje_asrs(respuestas_alumno, plantilla_seccion):
    """Calcula el puntaje del tamizaje ASRS-V1.1."""
    puntaje = 0
    opciones_sintomaticas = ["A menudo", "Muy a menudo"]
    
    for pregunta in plantilla_seccion['preguntas']:
        id_pregunta = pregunta['id']
        respuesta = respuestas_alumno.get(id_pregunta, "")
        if any(op in respuesta for op in opciones_sintomaticas):
            puntaje += 1
            
    interpretacion = "Sus síntomas NO son consistentes con TDAH del adulto."
    if puntaje >= 4:
        interpretacion = "Sus síntomas pueden ser consistentes con TDAH del adulto. Se requiere valoración integral."
        
    return puntaje, interpretacion

def calcular_puntaje_vinegrad(respuestas_alumno, plantilla_seccion):
    """Calcula el puntaje del tamizaje de dislexia de Vinegrad."""
    puntaje = 0
    opcion_afirmativa = "Sí"
    
    for pregunta in plantilla_seccion['preguntas']:
        id_pregunta = pregunta['id']
        respuesta = respuestas_alumno.get(id_pregunta, "")
        if respuesta.strip() == opcion_afirmativa:
            puntaje += 1
            
    interpretacion = "SIN riesgo de dificultades lectoras."
    if puntaje >= 9:
        interpretacion = "CON riesgo de dificultades lectoras. Se sugiere evaluación especializada."
        
    return puntaje, interpretacion

def crear_pdf(datos_alumno, plantilla, ruta_salida):
    """Genera un archivo PDF para un único alumno, incluyendo la página de resultados."""
    datos_alumno_str = {key: str(value) for key, value in datos_alumno.items()}

    # --- CAMBIO 2: Usamos la ruta completa 'units.cm' ---
    doc = SimpleDocTemplate(ruta_salida,
                            pagesize='letter',
                            rightMargin=2*units.cm, leftMargin=2*units.cm,
                            topMargin=2*units.cm, bottomMargin=2*units.cm)
    
    story = []
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Title', parent=styles['h1'], fontName='Helvetica-Bold', fontSize=14, alignment=TA_CENTER, spaceAfter=20))
    styles.add(ParagraphStyle(name='SectionTitle', parent=styles['h2'], fontName='Helvetica-Bold', fontSize=12, textColor=navy, spaceBefore=12, spaceAfter=6))
    styles.add(ParagraphStyle(name='Question', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceBefore=8))
    styles.add(ParagraphStyle(name='Answer', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leftIndent=1*units.cm, textColor=black))
    styles.add(ParagraphStyle(name='Result', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, textColor=crimson, spaceBefore=12, leftIndent=1*units.cm))

    story.append(Paragraph(plantilla['titulo_principal'], styles['Title']))

    seccion_aq10, seccion_asrs, seccion_vinegrad = None, None, None

    for seccion in plantilla['secciones']:
        story.append(Paragraph(seccion['titulo'], styles['SectionTitle']))
        
        if "Espectro Autista" in seccion['titulo']: seccion_aq10 = seccion
        if "Cribado del Adulto" in seccion['titulo']: seccion_asrs = seccion
        if "Dislexia en Edad Adulta" in seccion['titulo']: seccion_vinegrad = seccion
            
        for pregunta in seccion['preguntas']:
            respuesta = datos_alumno_str.get(pregunta['id'], "Sin respuesta")
            story.append(Paragraph(pregunta['texto'], styles['Question']))
            story.append(Paragraph(respuesta, styles['Answer']))

    story.append(PageBreak())
    story.append(Paragraph("Resultados de Tamizajes", styles['SectionTitle']))
    
    if seccion_aq10:
        puntaje, interpretacion = calcular_puntaje_aq10(datos_alumno_str, seccion_aq10)
        story.append(Paragraph("Tamizaje de Trastornos del Espectro Autista (AQ-10):", styles['Question']))
        story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 10", styles['Answer']))
        story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))
        story.append(Spacer(1, 0.5*units.cm))

    if seccion_asrs:
        puntaje, interpretacion = calcular_puntaje_asrs(datos_alumno_str, seccion_asrs)
        story.append(Paragraph("Cuestionario de Cribado del Adulto (ASRS-V1.1):", styles['Question']))
        story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 6", styles['Answer']))
        story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))
        story.append(Spacer(1, 0.5*units.cm))

    if seccion_vinegrad:
        puntaje, interpretacion = calcular_puntaje_vinegrad(datos_alumno_str, seccion_vinegrad)
        story.append(Paragraph("Lista de Control para la Dislexia en Edad Adulta (Vinegrad):", styles['Question']))
        story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 20 respuestas afirmativas", styles['Answer']))
        story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))

    doc.build(story)

def main():
    """Función principal que orquesta todo el proceso."""
    print("🚀 Iniciando el generador de reportes en PDF...")
    
    ruta_excel = seleccionar_archivo_excel()
    if not ruta_excel:
        print("❌ No se seleccionó ningún archivo. Saliendo del programa.")
        return

    directorio_salida = seleccionar_directorio_salida()
    if not directorio_salida:
        print("❌ No se seleccionó ninguna carpeta de salida. Saliendo del programa.")
        return

    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            plantilla = json.load(f)
        
        df = pd.read_excel(ruta_excel)

    except FileNotFoundError:
        print("🚨 Error: No se encontró el archivo 'config.json'. Asegúrate de que esté en la misma carpeta que el script.")
        return
    except Exception as e:
        print(f"🚨 Error inesperado al leer los archivos: {e}")
        return

    print(f"\n✅ Archivos cargados. Se encontraron {len(df)} registros para procesar.")
    
    for index, alumno in df.iterrows():
        datos_alumno_dict = alumno.to_dict()
        clave_unica = str(datos_alumno_dict.get('¿Cuál es tu Clave Única?', f'registro_{index + 1}'))
        
        if clave_unica.endswith('.0'):
            clave_unica = clave_unica[:-2]

        nombre_archivo = f"{clave_unica}.pdf"
        ruta_completa_salida = os.path.join(directorio_salida, nombre_archivo)
        
        print(f"📄 Generando PDF para el alumno con Clave Única: {clave_unica}...")
        
        try:
            crear_pdf(datos_alumno_dict, plantilla, ruta_completa_salida)
        except Exception as e:
            print(f"  ❗️ ERROR DETALLADO al generar el PDF para {clave_unica}:")
            traceback.print_exc()

    print("\n🎉 ¡Proceso completado! Todos los PDFs han sido generados en la carpeta seleccionada.")

if __name__ == "__main__":
    main()