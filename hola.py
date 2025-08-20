import pandas as pd
import json
import os
import traceback
from tkinter import Tk, filedialog
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import units
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
from reportlab.lib.colors import navy, black, crimson, grey, HexColor
from datetime import datetime

# --- Clase para manejar Encabezado y Pie de Página ---
class HeaderFooterCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.student_name = kwargs.pop('student_name', 'N/A')
        self.student_cu = kwargs.pop('student_cu', 'N/A')
        self.header_text = kwargs.pop('header_text', 'Reporte')
        canvas.Canvas.__init__(self, *args, **kwargs)

    def save(self):
        self.setFont('Helvetica', 8)
        self.setFillColor(grey)
        # Encabezado (Posición Y ajustada para subirlo)
        header_y_position = self._pagesize[1] - 1.0 * units.cm
        self.drawString(2.5 * units.cm, header_y_position, self.header_text)
        self.drawRightString(self._pagesize[0] - 2.5 * units.cm, header_y_position, f"{self.student_name} ({self.student_cu})")
        # Pie de página
        self.drawRightString(self._pagesize[0] - 2.5 * units.cm, 1.5 * units.cm, f"Página {self.getPageNumber()}")
        canvas.Canvas.save(self)

# --- Funciones de Cálculo ---
def calcular_edad(fecha_nacimiento, fecha_inicio):
    try:
        fn = pd.to_datetime(fecha_nacimiento)
        fi = pd.to_datetime(fecha_inicio)
        edad = fi.year - fn.year - ((fi.month, fi.day) < (fn.month, fn.day))
        return str(edad)
    except Exception:
        return "N/A"

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
    if puntaje >= 6: interpretacion = "Puntaje sugestivo. Se recomienda valoración especializada."
    return puntaje, interpretacion

def calcular_puntaje_asrs(respuestas_alumno, questions_config):
    puntaje = 0
    opciones_sintomaticas = ["a menudo", "muy a menudo"]
    for pregunta in questions_config:
        respuesta = respuestas_alumno.get(pregunta, "").lower()
        if any(op in respuesta for op in opciones_sintomaticas): puntaje += 1
    interpretacion = "Síntomas no consistentes con TDAH del adulto."
    if puntaje >= 4: interpretacion = "Síntomas pueden ser consistentes con TDAH del adulto. Se requiere valoración integral."
    return puntaje, interpretacion

def calcular_puntaje_vinegrad(respuestas_alumno, questions_config):
    puntaje = 0
    for pregunta in questions_config:
        if respuestas_alumno.get(pregunta, "").strip().lower() == "sí": puntaje += 1
    interpretacion = "SIN riesgo de dificultades lectoras."
    if puntaje >= 9: interpretacion = "CON riesgo de dificultades lectoras. Se sugiere evaluación especializada."
    return puntaje, interpretacion

# --- Función de Creación de PDF ---
def crear_pdf(datos_alumno, config, ruta_salida, styles):
    student_cu = str(datos_alumno.get('clave_unica', 'N/A'))
    student_name = str(datos_alumno.get('nombre_completo', 'N/A'))
    
    # Margen superior ajustado para subir el contenido
    doc = SimpleDocTemplate(ruta_salida, pagesize=letter, topMargin=1.5*units.cm)
    story = []
    
    story.append(Paragraph("CÉDULA DE TAMIZAJE UNIVERSITARIO", styles['Title']))

    # --- Sección 1: Datos Personales y Contexto ---
    story.append(Paragraph("1. Datos Personales y Contexto", styles['SectionTitle']))
    
    datos_personales = [
        ('Nombre Completo:', str(datos_alumno.get('nombre_completo', ''))),
        ('Clave Única:', str(datos_alumno.get('clave_unica', ''))),
        ('Fecha de Nacimiento:', str(pd.to_datetime(datos_alumno.get('fecha_nacimiento')).strftime('%Y-%m-%d') if pd.notna(datos_alumno.get('fecha_nacimiento')) else '')),
        ('Edad (calculada):', str(datos_alumno.get('edad_calculada', ''))),
        ('Género:', str(datos_alumno.get('genero', ''))),
        ('Entidad Académica:', str(datos_alumno.get('entidad', ''))),
        ('Carrera:', str(datos_alumno.get('carrera', ''))),
        ('¿Se percibe como indígena?:', str(datos_alumno.get('grupo_indigena', ''))),
        ('¿Tiene diversidad funcional?:', str(datos_alumno.get('tiene_diversidad', ''))),
        ('¿Tiene diagnóstico médico/psic?:', str(datos_alumno.get('diagnostico_medico', ''))),
        ('¿Ha recibido tratamiento farmacológico?:', str(datos_alumno.get('tratamiento_farma', ''))),
        ('¿Ha recibido tratamiento psicológico?:', str(datos_alumno.get('tratamiento_psico', '')))
    ]
    tabla_personales = Table(datos_personales, colWidths=[6 * units.cm, 10 * units.cm])
    tabla_personales.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, grey)
    ]))
    story.append(tabla_personales)
    story.append(Spacer(1, 0.5*units.cm))

    # --- Sección 2: Antecedentes y Desarrollo en Tres Columnas ---
    
    def preparar_columna(titulo, keys, datos, estilos, condicion_especial=False):
        contenido = [Paragraph(f"<b>{titulo}</b>", estilos['ColumnHeader'])]
        hallazgos = 0
        for key in keys:
            pregunta = next((q for q, k in config['column_mapping'].items() if k == key), key)
            respuesta = str(datos.get(key, "")).lower()
            
            mostrar = False
            if not condicion_especial:
                if 'sí' in respuesta: mostrar = True
            else: 
                if 'sí' in respuesta or 'no sé' in respuesta or 'no recuerdo' in respuesta:
                    mostrar = True
            
            if mostrar:
                contenido.append(Paragraph(f"• {pregunta}", estilos['ColumnText']))
                hallazgos += 1

        if hallazgos == 0:
            contenido.append(Paragraph("Sin hallazgos reportados.", estilos['ColumnText']))
        return contenido

    med_keys = [s['keys'] for s in config['sections'] if 'Médicos Personales' in s['title']][0]
    heredo_keys = [s['keys'] for s in config['sections'] if 'Heredofamiliares' in s['title']][0]
    desarrollo_keys = [s['keys'] for s in config['sections'] if 'Desarrollo' in s['title']][0]

    col_medicos = preparar_columna("Antecedentes Médicos", med_keys, datos_alumno, styles)
    col_heredo = preparar_columna("Antecedentes Heredofamiliares", heredo_keys, datos_alumno, styles)
    col_desarrollo = preparar_columna("Historial de Desarrollo", desarrollo_keys, datos_alumno, styles, condicion_especial=True)

    tabla_columnas_data = [[col_medicos, col_heredo, col_desarrollo]]
    tabla_columnas = Table(tabla_columnas_data, colWidths=[5.5*units.cm, 5.5*units.cm, 5.5*units.cm])
    tabla_columnas.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, grey),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(tabla_columnas)

    # --- Sección 3: Resultados de Instrumentos de Tamizaje ---
    story.append(Paragraph("3. Resultados de Instrumentos de Tamizaje", styles['SectionTitle']))
    
    puntaje_aq10, interpretacion_aq10 = calcular_puntaje_aq10(datos_alumno, config['evaluations']['aq10']['questions'])
    story.append(Paragraph(config['evaluations']['aq10']['title'], styles['SubSectionTitle']))
    story.append(Paragraph(f"<b>Puntaje Obtenido:</b> {puntaje_aq10} de 10", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion_aq10}", styles['Result']))
    story.append(Spacer(1, 0.5*units.cm))

    puntaje_asrs, interpretacion_asrs = calcular_puntaje_asrs(datos_alumno, config['evaluations']['asrs']['questions'])
    story.append(Paragraph(config['evaluations']['asrs']['title'], styles['SubSectionTitle']))
    story.append(Paragraph(f"<b>Puntaje Obtenido:</b> {puntaje_asrs} de 6", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion_asrs}", styles['Result']))
    story.append(Spacer(1, 0.5*units.cm))

    puntaje_vinegrad, interpretacion_vinegrad = calcular_puntaje_vinegrad(datos_alumno, config['evaluations']['vinegrad']['questions'])
    story.append(Paragraph(config['evaluations']['vinegrad']['title'], styles['SubSectionTitle']))
    story.append(Paragraph(f"<b>Puntaje Obtenido:</b> {puntaje_vinegrad} de 20 respuestas afirmativas", styles['Answer']))
    story.append(Paragraph(f"<b>Interpretación:</b> {interpretacion_vinegrad}", styles['Result']))
    
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
    
    columnas_entidad = [col for col in df.columns if 'Facultad' in col or 'Coordinación' in col or 'Unidad' in col]
    if columnas_entidad:
        df['carrera'] = df[columnas_entidad].bfill(axis=1).iloc[:, 0]
    else:
        df['carrera'] = 'No especificada'
        
    df['edad_calculada'] = df.apply(lambda row: calcular_edad(row['fecha_nacimiento'], datetime.now()), axis=1)

    # --- CREACIÓN DE ESTILOS ---
    styles = getSampleStyleSheet()
    
    styles['Title'].fontName = 'Helvetica-Bold'
    styles['Title'].fontSize = 16
    styles['Title'].alignment = TA_CENTER
    styles['Title'].spaceAfter = 20
    styles['Title'].textColor = navy
    
    styles.add(ParagraphStyle(name='SectionTitle', parent=styles['h2'], fontName='Helvetica-Bold', fontSize=13, textColor=navy, spaceBefore=12, spaceAfter=6, borderPadding=2, borderColor=navy, borderBottomWidth=0.5))
    styles.add(ParagraphStyle(name='SubSectionTitle', parent=styles['h3'], fontName='Helvetica-Bold', fontSize=11, textColor=navy, spaceBefore=10, spaceAfter=4))
    styles.add(ParagraphStyle(name='Question', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, spaceBefore=8))
    styles.add(ParagraphStyle(name='Answer', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leftIndent=1*units.cm, textColor=black, spaceBefore=4))
    styles.add(ParagraphStyle(name='Result', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=crimson, spaceBefore=4, leftIndent=1*units.cm))
    styles.add(ParagraphStyle(name='ColumnHeader', fontName='Helvetica-Bold', fontSize=10, alignment=TA_CENTER, spaceAfter=6))
    styles.add(ParagraphStyle(name='ColumnText', parent=styles['Normal'], fontSize=9, spaceBefore=2))

    
    for _, row in df.iterrows():
        datos_alumno_dict = row.to_dict()
        clave_unica = str(datos_alumno_dict.get('clave_unica', f'registro_{_ + 1}'))
        if clave_unica.endswith('.0'): clave_unica = clave_unica[:-2]
        
        nombre_archivo = f"{clave_unica}.pdf"
        ruta_completa_salida = os.path.join(output_folder, nombre_archivo)
        
        print(f"📄 Generando PDF para Clave Única: {clave_unica}...")
        try:
            crear_pdf(datos_alumno_dict, config, ruta_completa_salida, styles)
        except Exception:
            print(f"  ❗️ ERROR DETALLADO al generar el PDF para {clave_unica}:")
            traceback.print_exc()
            
    print("\n🎉 ¡Proceso completado!")

if __name__ == "__main__":
    main()