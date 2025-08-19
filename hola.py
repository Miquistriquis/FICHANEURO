import pandas as pd
import json
import os
import traceback
from tkinter import Tk, filedialog
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import units
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.colors import navy, black, crimson, grey

# --- Clase para manejar Encabezado y Pie de Página ---
class HeaderFooterCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.student_name = kwargs.pop('student_name', 'N/A')
        self.student_cu = kwargs.pop('student_cu', 'N/A')
        self.header_text = kwargs.pop('header_text', 'Reporte')
        canvas.Canvas.__init__(self, *args, **kwargs)

    def save(self):
        self.setFont('Helvetica', 8)
        # Encabezado
        self.drawString(2.5 * units.cm, self._pagesize[1] - 1.5 * units.cm, self.header_text)
        self.drawRightString(self._pagesize[0] - 2.5 * units.cm, self._pagesize[1] - 1.5 * units.cm, f"{self.student_name} ({self.student_cu})")
        # Pie de página
        self.drawRightString(self._pagesize[0] - 2.5 * units.cm, 1.5 * units.cm, f"Página {self.getPageNumber()}")
        canvas.Canvas.save(self)

# --- Funciones de Cálculo de Puntajes ---
def calcular_puntaje_aq10(respuestas_alumno, questions_config):
    puntaje = 0
    items_positivos = [1, 7, 8, 10]
    opciones_acuerdo = ["(3)", "de acuerdo"]
    opciones_desacuerdo = ["(1)", "(2)", "desacuerdo"]
    for i, pregunta in enumerate(questions_config, 1):
        respuesta = respuestas_alumno.get(pregunta, "").lower()
        if i in items_positivos:
            if any(op in respuesta for op in opciones_acuerdo): puntaje += 1
        else:
            if any(op in respuesta for op in opciones_desacuerdo): puntaje += 1
    interpretacion = "Puntaje dentro del rango esperado."
    if puntaje >= 6: interpretacion = "Puntaje sugestivo. Se recomienda una valoración más profunda por un especialista."
    return puntaje, interpretacion

def calcular_puntaje_asrs(respuestas_alumno, questions_config):
    puntaje = 0
    opciones_sintomaticas = ["a menudo", "muy a menudo"]
    for pregunta in questions_config:
        respuesta = respuestas_alumno.get(pregunta, "").lower()
        if any(op in respuesta for op in opciones_sintomaticas): puntaje += 1
    interpretacion = "Sus síntomas NO son consistentes con TDAH del adulto."
    if puntaje >= 4: interpretacion = "Sus síntomas pueden ser consistentes con TDAH del adulto. Se requiere valoración integral."
    return puntaje, interpretacion

def calcular_puntaje_vinegrad(respuestas_alumno, questions_config):
    puntaje = 0
    for pregunta in questions_config:
        if respuestas_alumno.get(pregunta, "").strip().lower() == "sí": puntaje += 1
    interpretacion = "SIN riesgo de dificultades lectoras."
    if puntaje >= 9: interpretacion = "CON riesgo de dificultades lectoras. Se sugiere evaluación especializada."
    return puntaje, interpretacion

# --- Función Principal de Creación de PDF ---
def crear_pdf(datos_alumno, config, ruta_salida, styles):
    student_cu = str(datos_alumno.get('clave_unica', 'N/A'))
    student_name = str(datos_alumno.get('nombre_completo', 'N/A'))
    
    doc = SimpleDocTemplate(ruta_salida, pagesize=letter)
    story = []
    
    story.append(Paragraph("CÉDULA DE TAMIZAJE UNIVERSITARIO", styles['Title']))

    # --- Secciones de Datos ---
    for section_config in config['sections']:
        story.append(Paragraph(section_config['title'], styles['SectionTitle']))
        for key in section_config['keys']:
            pregunta_original = next((q for q, k in config['column_mapping'].items() if k == key), key.replace('_', ' ').capitalize())
            respuesta = str(datos_alumno.get(key, "Sin respuesta"))
            if respuesta.lower() not in ['nan', '']:
                story.append(Paragraph(pregunta_original, styles['Question']))
                story.append(Paragraph(respuesta, styles['Answer']))
    
    # --- Secciones de Evaluaciones (Respuestas) ---
    for eval_key, eval_config in config['evaluations'].items():
        story.append(Paragraph(f"Respuestas: {eval_config['title']}", styles['SectionTitle']))
        for i, pregunta in enumerate(eval_config['questions'], 1):
            respuesta = str(datos_alumno.get(pregunta, "Sin respuesta"))
            if respuesta.lower() not in ['nan', '']:
                story.append(Paragraph(f"{i}. {pregunta[:80]}...", styles['Question'])) # Acortar pregunta para legibilidad
                story.append(Paragraph(respuesta, styles['Answer']))

    # --- Página de Resultados ---
    story.append(PageBreak())
    story.append(Paragraph("Resultados de Tamizajes", styles['Title']))
    
    # AQ-10
    puntaje, interpretacion = calcular_puntaje_aq10(datos_alumno, config['evaluations']['aq10']['questions'])
    story.append(Paragraph(config['evaluations']['aq10']['title'], styles['SectionTitle']))
    story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 10", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))
    story.append(Spacer(1, 0.8*units.cm))

    # ASRS
    puntaje, interpretacion = calcular_puntaje_asrs(datos_alumno, config['evaluations']['asrs']['questions'])
    story.append(Paragraph(config['evaluations']['asrs']['title'], styles['SectionTitle']))
    story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 6", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))
    story.append(Spacer(1, 0.8*units.cm))

    # Vinegrad
    puntaje, interpretacion = calcular_puntaje_vinegrad(datos_alumno, config['evaluations']['vinegrad']['questions'])
    story.append(Paragraph(config['evaluations']['vinegrad']['title'], styles['SectionTitle']))
    story.append(Paragraph(f"Puntaje Obtenido: {puntaje} de 20 respuestas afirmativas", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion}", styles['Result']))
    
    header_footer_args = {
        'student_name': student_name,
        'student_cu': student_cu,
        'header_text': config['pdf_styles']['header_text']
    }
    doc.build(story, canvasmaker=lambda *args, **kwargs: HeaderFooterCanvas(*args, **kwargs, **header_footer_args))

# --- Función Principal de Ejecución ---
def main():
    print("🚀 Iniciando el generador de reportes en PDF...")
    root = Tk()
    root.withdraw()
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        print("🚨 Error: No se encontró 'config.json'. Asegúrate de que esté en la misma carpeta.")
        return

    excel_file = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not excel_file:
        print("❌ No se seleccionó ningún archivo. Saliendo.")
        return
        
    output_folder = filedialog.askdirectory(title="Selecciona la carpeta para guardar los PDFs")
    if not output_folder:
        print("❌ No se seleccionó ninguna carpeta de salida. Saliendo.")
        return

    try:
        df = pd.read_excel(excel_file)
        df.rename(columns=config['column_mapping'], inplace=True)
    except Exception as e:
        print(f"🚨 Error inesperado al leer el archivo Excel: {e}")
        return

    print(f"\n✅ Archivos cargados. Se encontraron {len(df)} registros para procesar.")
    
    # --- Lógica para encontrar la carrera ---
    columnas_entidad = [col for col in df.columns if 'Facultad' in col or 'Coordinación' in col or 'Unidad' in col]
    if columnas_entidad:
        df['carrera'] = df[columnas_entidad].bfill(axis=1).iloc[:, 0]
    else:
        df['carrera'] = 'No especificada'

    # --- CREACIÓN DE ESTILOS (CORREGIDO) ---
    styles = getSampleStyleSheet()
    
    # Modificamos el estilo 'Title' que ya existe
    styles['Title'].fontName = 'Helvetica-Bold'
    styles['Title'].fontSize = 16
    styles['Title'].alignment = TA_CENTER
    styles['Title'].spaceAfter = 20
    styles['Title'].textColor = navy
    
    # Agregamos los nuevos estilos que no existen
    styles.add(ParagraphStyle(name='SectionTitle', parent=styles['h2'], fontName='Helvetica-Bold', fontSize=12, textColor=navy, spaceBefore=12, spaceAfter=6, borderPadding=2, borderColor=navy, borderBottomWidth=0.5))
    styles.add(ParagraphStyle(name='Question', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceBefore=8))
    styles.add(ParagraphStyle(name='Answer', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leftIndent=1*units.cm, textColor=black))
    styles.add(ParagraphStyle(name='Result', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, textColor=crimson, spaceBefore=12, leftIndent=1*units.cm))
    
    for _, row in df.iterrows():
        datos_alumno_dict = row.to_dict()
        clave_unica = str(datos_alumno_dict.get('clave_unica', f'registro_{_ + 1}'))
        if clave_unica.endswith('.0'): clave_unica = clave_unica[:-2]
        
        nombre_archivo = f"{clave_unica}.pdf"
        ruta_completa_salida = os.path.join(output_folder, nombre_archivo)
        
        print(f"📄 Generando PDF para Clave Única: {clave_unica}...")
        try:
            # Pasamos los estilos ya creados a la función
            crear_pdf(datos_alumno_dict, config, ruta_completa_salida, styles)
        except Exception:
            print(f"  ❗️ ERROR DETALLADO al generar el PDF para {clave_unica}:")
            traceback.print_exc()
            
    print("\n🎉 ¡Proceso completado!")

if __name__ == "__main__":
    main()