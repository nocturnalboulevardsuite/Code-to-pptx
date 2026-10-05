import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


def generar_presentacion_pptx() -> io.BytesIO:
    # Paleta de Colores
    DARK_BG = RGBColor(15, 7, 10)       # #0f070a - Fondo oscuro elegante
    LIGHT_BG = RGBColor(255, 248, 249)  # #fff8f9 - Fondo claro
    WHITE = RGBColor(255, 255, 255)
    DARK_TEXT = RGBColor(17, 24, 39)     # #111827 - Texto oscuro
    SUBTITLE_TEXT = RGBColor(100, 116, 139) # #64748b - Texto secundario
    RED_PRIMARY = RGBColor(225, 29, 72)  # #e11d48 - Rojo primario
    RED_CRIMSON = RGBColor(244, 63, 94)  # #f43f5e - Rojo carmesí
    RED_LIGHT = RGBColor(255, 228, 230)  # #ffe4e6 - Fondo sutil tarjeta
    CARD_DARK = RGBColor(30, 15, 22)     # Tarjeta oscura

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6]

    def set_slide_background(slide, color):
        """Establece el color de fondo de una diapositiva."""
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, badge_text, title_part1, title_part2, subtitle=""):
        """Agrega un encabezado consistente en diapositivas de fondo claro."""
        badge_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(4), Inches(0.4))
        tf_b = badge_box.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = badge_text.upper()
        p_b.font.size = Pt(11)
        p_b.font.bold = True
        p_b.font.color.rgb = RED_PRIMARY

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        
        run1 = p_t.add_run()
        run1.text = title_part1 + " "
        run1.font.size = Pt(26)
        run1.font.bold = True
        run1.font.color.rgb = DARK_TEXT
        
        run2 = p_t.add_run()
        run2.text = title_part2
        run2.font.size = Pt(26)
        run2.font.bold = True
        run2.font.color.rgb = RED_PRIMARY

        if subtitle:
            p_sub = tf_t.add_paragraph()
            p_sub.text = subtitle
            p_sub.font.size = Pt(13)
            p_sub.font.color.rgb = SUBTITLE_TEXT

    # --------------------------------------------------------------------------
    # Diapositiva 1: Portada
    # --------------------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide1, DARK_BG)

    title_box = slide1.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.0))
    tf1 = title_box.text_frame
    tf1.word_wrap = True

    p_badge = tf1.paragraphs[0]
    p_badge.alignment = PP_ALIGN.CENTER
    p_badge.text = "ESTRATEGIA DE MODERNIZACIÓN & LOGÍSTICA"
    p_badge.font.size = Pt(12)
    p_badge.font.bold = True
    p_badge.font.color.rgb = RED_CRIMSON

    p_main = tf1.add_paragraph()
    p_main.alignment = PP_ALIGN.CENTER
    r1 = p_main.add_run()
    r1.text = "Transformación Digital y "
    r1.font.size = Pt(38)
    r1.font.bold = True
    r1.font.color.rgb = WHITE

    r2 = p_main.add_run()
    r2.text = "Agentes de IA"
    r2.font.size = Pt(38)
    r2.font.bold = True
    r2.font.color.rgb = RED_CRIMSON

    r3 = p_main.add_run()
    r3.text = "\nen Abastecimiento"
    r3.font.size = Pt(38)
    r3.font.bold = True
    r3.font.color.rgb = WHITE

    p_sub = tf1.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = "\nOptimizando el ciclo de compras con diagnóstico continuo y orquestación inteligente"
    p_sub.font.size = Pt(16)
    p_sub.font.color.rgb = RGBColor(253, 164, 175)

    # --------------------------------------------------------------------------
    # Diapositiva 2: Flujo Operativo
    # --------------------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide2, LIGHT_BG)
    add_header(slide2, "Flujo Operativo", "Línea de Tiempo del", "Ciclo de Compra", "Etapas clave sujetas a trazabilidad y automatización")

    steps = [
        ("1", "SOLPED", "Creación y liberación de la solicitud interna de compra."),
        ("2", "Scouting", "Búsqueda, cotización con proveedores y cuadro comparativo."),
        ("3", "Adjudicación", "Negociación, selección del proveedor y emisión de OC."),
        ("4", "Logística", "Despacho, recepción física de bienes y pago a proveedor.")
    ]

    col_width = Inches(2.6)
    start_x = Inches(0.8)
    gap = Inches(0.3)

    for i, (num, title, desc) in enumerate(steps):
        x_pos = start_x + i * (col_width + gap)
        
        shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x_pos, Inches(2.3), col_width, Inches(4.2))
        shape.fill.solid()
        shape.fill.fore_color.rgb = WHITE
        shape.line.color.rgb = RED_LIGHT
        
        tf = shape.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.text = num
        p0.font.size = Pt(28)
        p0.font.bold = True
        p0.font.color.rgb = RED_PRIMARY
        
        p1 = tf.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        p1.text = title
        p1.font.size = Pt(18)
        p1.font.bold = True
        p1.font.color.rgb = DARK_TEXT
        
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.text = f"\n{desc}"
        p2.font.size = Pt(13)
        p2.font.color.rgb = SUBTITLE_TEXT

    # --------------------------------------------------------------------------
    # Diapositiva 3: Conceptos Clave
    # --------------------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide3, LIGHT_BG)
    add_header(slide3, "Conceptos Clave", "Glosario Explicativo de", "Indicadores", "Términos fundamentales para medir la eficiencia del área y la gestión de compras.")

    glossary = [
        ("OTIF", "Nivel de Servicio", "On-Time In-Full (A Tiempo y Completo)", "Mide el % de pedidos que los proveedores entregan a tiempo y completos en las cantidades requeridas.", "Impacto: Evalúa el cumplimiento logístico directo del proveedor."),
        ("TAT", "Tiempo de Respuesta", "Turnaround Time (Tiempo de Ciclo Total)", "Total de días transcurridos desde que se emite una SOLPED hasta la emisión formal de la Orden de Compra (OC).", "Impacto: Detecta cuellos de botella en negociación y aprobación."),
        ("Dx Compradores", "Panel Analítico", "Diagnóstico Operativo de Compradores", "Dashboard para monitorear en tiempo real la carga de trabajo individual, antigüedad de solicitudes y balance.", "Impacto: Permite redistribuir carga y evitar retrasos en backlog."),
        ("SOLPED", "Requerimiento", "Solicitud de Pedido", "Petición formal redactada internamente por un área usuaria en el sistema (SAP/ERP) para requerir bienes o servicios.", "Impacto: Punto inicial que desencadena la gestión de compra.")
    ]

    positions = [
        (Inches(0.8), Inches(1.8)),
        (Inches(6.8), Inches(1.8)),
        (Inches(0.8), Inches(4.5)),
        (Inches(6.8), Inches(4.5))
    ]

    for idx, (term, tag, full_name, text, impact) in enumerate(glossary):
        x, y = positions[idx]
        card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.4))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RED_PRIMARY

        tf = card.text_frame
        tf.word_wrap = True
        
        p_title = tf.paragraphs[0]
        r_t = p_title.add_run()
        r_t.text = term + "  "
        r_t.font.size = Pt(20)
        r_t.font.bold = True
        r_t.font.color.rgb = RED_PRIMARY
        
        r_tag = p_title.add_run()
        r_tag.text = f"[{tag}]"
        r_tag.font.size = Pt(11)
        r_tag.font.bold = True
        r_tag.font.color.rgb = SUBTITLE_TEXT
        
        p_sub = tf.add_paragraph()
        p_sub.text = full_name.upper()
        p_sub.font.size = Pt(10)
        p_sub.font.bold = True
        p_sub.font.color.rgb = RED_CRIMSON

        p_desc = tf.add_paragraph()
        p_desc.text = text
        p_desc.font.size = Pt(12)
        p_desc.font.color.rgb = DARK_TEXT

        p_imp = tf.add_paragraph()
        p_imp.text = impact
        p_imp.font.size = Pt(11)
        p_imp.font.bold = True
        p_imp.font.color.rgb = RED_PRIMARY

    # --------------------------------------------------------------------------
    # Diapositiva 4: Mejora de Procesos
    # --------------------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide4, LIGHT_BG)
    add_header(slide4, "Mejora de Procesos", "Avances Analíticos:", "Control SLA y Rediseño", "Evolución hacia un diagnóstico operativo continuo de compradores.")

    features = [
        ("Trazabilidad Punta a Punta (TAT / OTIF)", "Medición exacta de tiempos desde la emisión de SOLPED hasta la entrega física del proveedor."),
        ("Identificación de Cuellos de Botella", "Detección temprana de bloqueos en validaciones técnicas, presupuestarias o aprobaciones."),
        ("Evaluación Objetiva de Proveedores", "Matriz de rendimiento logístico basada en datos reales para negociaciones estratégicas.")
    ]

    for i, (f_title, f_desc) in enumerate(features):
        card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0 + i * 1.6), Inches(11.7), Inches(1.3))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RED_LIGHT
        
        tf = card.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r_head = p.add_run()
        r_head.text = f_title + "\n"
        r_head.font.size = Pt(16)
        r_head.font.bold = True
        r_head.font.color.rgb = DARK_TEXT
        
        r_body = p.add_run()
        r_body.text = f_desc
        r_body.font.size = Pt(13)
        r_body.font.color.rgb = SUBTITLE_TEXT

    # --------------------------------------------------------------------------
    # Diapositiva 5: Evolución Tecnológica
    # --------------------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide5, LIGHT_BG)
    add_header(slide5, "Evolución Tecnológica", "Optimización del Módulo", "Dx Compradores", "Transformación de datos estáticos en monitoreo diario y accionable.")

    dx_cards = [
        ("Cómputo Continuo (Regla 1+ Días)", "Cálculo dinámico en tiempo real que mide la antigüedad exacta de cada solicitud abierta diariamente.\n\n• Elimina la medición semanal congelada.\n• Visibilidad inmediata de solicitudes con riesgo de atraso."),
        ("Mejoras de Usabilidad Operativa", "Herramientas diseñadas para agilizar la labor diaria del equipo de abastecimiento.\n\n• Persistencia de notas por SOLPED.\n• Filtro masivo por lote de IDs.\n• Clasificación automática de orígenes (SAP/Ariba).")
    ]

    for i, (title, body) in enumerate(dx_cards):
        card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 5.9), Inches(2.0), Inches(5.7), Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RED_PRIMARY

        tf = card.text_frame
        tf.word_wrap = True
        
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_TEXT

        p_b = tf.add_paragraph()
        p_b.text = f"\n{body}"
        p_b.font.size = Pt(13)
        p_b.font.color.rgb = SUBTITLE_TEXT

    # --------------------------------------------------------------------------
    # Diapositiva 6: Próxima Fase de Automatización
    # --------------------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide6, LIGHT_BG)
    add_header(slide6, "Próxima Fase de Automatización", "Arquitectura", "Multiagente de IA", "Automatización cognitiva para tareas repetitivas de cotización y análisis.")

    agents = [
        ("Agente Orquestador (Triage)", "Estrategia y Clasificación", "Interpreta el requerimiento de la SOLPED. Clasifica el tipo de compra y evalúa si aplica contrato marco o cotización estándar."),
        ("Agentes Especialistas", "Extracción & Comparación", "Envían RFQs, procesan PDFs recibidos, extraen precios y plazos, y construyen la matriz comparativa automáticamente.")
    ]

    for i, (title, tag, desc) in enumerate(agents):
        card = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 5.9), Inches(2.0), Inches(5.7), Inches(3.2))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RED_PRIMARY

        tf = card.text_frame
        tf.word_wrap = True
        
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_TEXT

        p_tag = tf.add_paragraph()
        p_tag.text = tag
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = RED_PRIMARY

        p_d = tf.add_paragraph()
        p_d.text = f"\n{desc}"
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = SUBTITLE_TEXT

    # Banner Human-in-the-loop
    banner = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.4), Inches(11.6), Inches(1.3))
    banner.fill.solid()
    banner.fill.fore_color.rgb = DARK_BG
    banner.line.color.rgb = RED_PRIMARY

    tf_b = banner.text_frame
    tf_b.word_wrap = True
    p_bt = tf_b.paragraphs[0]
    p_bt.text = "Control Humano (Human-in-the-Loop)"
    p_bt.font.size = Pt(15)
    p_bt.font.bold = True
    p_bt.font.color.rgb = WHITE

    p_bd = tf_b.add_paragraph()
    p_bd.text = "El comprador mantiene el control final: lidera la negociación estratégica con proveedores y aprueba la adjudicación."
    p_bd.font.size = Pt(12)
    p_bd.font.color.rgb = RED_LIGHT

    # --------------------------------------------------------------------------
    # Diapositiva 7: Stack de Implementación
    # --------------------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide7, LIGHT_BG)
    add_header(slide7, "Stack de Implementación", "Requerimientos Tecnológicos", "(Habilitación TI)", "Infraestructura base requerida para desplegar los Agentes de IA.")

    tech_stack = [
        ("n8n", "Orquestación", "Plataforma de Flujos", "Conecta los sistemas internos (SAP, correos, DBs) con los modelos de IA de forma segura."),
        ("Claude API", "Motor IA", "Inteligencia Cognitiva", "Modelo de precisión para lectura de documentos no estructurados (PDFs, ofertas) y razonamiento."),
        ("Claude Pro", "Prototipado", "Laboratorio de Prompts", "Entorno para diseño rápido de prompts, validación de reglas de negocio y pruebas continuas.")
    ]

    for i, (title, tag, sub, desc) in enumerate(tech_stack):
        card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 3.9), Inches(2.0), Inches(3.7), Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RED_PRIMARY

        tf = card.text_frame
        tf.word_wrap = True
        
        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(22)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_TEXT

        p_tag = tf.add_paragraph()
        p_tag.text = f"[{tag}]"
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = RED_PRIMARY

        p_sub = tf.add_paragraph()
        p_sub.text = sub
        p_sub.font.size = Pt(13)
        p_sub.font.bold = True
        p_sub.font.color.rgb = SUBTITLE_TEXT

        p_d = tf.add_paragraph()
        p_d.text = f"\n{desc}"
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = DARK_TEXT

    # --------------------------------------------------------------------------
    # Diapositiva 8: Hoja de Ruta
    # --------------------------------------------------------------------------
    slide8 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(slide8, DARK_BG)

    title_box8 = slide8.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11.33), Inches(1.2))
    tf8 = title_box8.text_frame
    tf8.word_wrap = True

    p_h8 = tf8.paragraphs[0]
    p_h8.alignment = PP_ALIGN.CENTER
    r_h1 = p_h8.add_run()
    r_h1.text = "Hoja de Ruta: "
    r_h1.font.size = Pt(32)
    r_h1.font.bold = True
    r_h1.font.color.rgb = WHITE

    r_h2 = p_h8.add_run()
    r_h2.text = "Agentes de IA"
    r_h2.font.size = Pt(32)
    r_h2.font.bold = True
    r_h2.font.color.rgb = RED_CRIMSON

    phases = [
        ("FASE 1", "Infraestructura", "Despliegue de n8n e integración segura con la API de Claude."),
        ("FASE 2", "Laboratorio & Prompts", "Diseño de reglas de negocio y entrenamiento con Claude Pro."),
        ("FASE 3", "Piloto MVP", "Primer ciclo de cotización autónoma con supervisión del comprador.")
    ]

    for i, (phase, p_title, p_desc) in enumerate(phases):
        card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 3.9), Inches(2.5), Inches(3.7), Inches(4.0))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_DARK
        card.line.color.rgb = RED_PRIMARY

        tf = card.text_frame
        tf.word_wrap = True
        
        p_ph = tf.paragraphs[0]
        p_ph.text = phase
        p_ph.font.size = Pt(11)
        p_ph.font.bold = True
        p_ph.font.color.rgb = RED_CRIMSON

        p_pt = tf.add_paragraph()
        p_pt.text = p_title
        p_pt.font.size = Pt(18)
        p_pt.font.bold = True
        p_pt.font.color.rgb = WHITE

        p_pd = tf.add_paragraph()
        p_pd.text = f"\n{p_desc}"
        p_pd.font.size = Pt(13)
        p_pd.font.color.rgb = RED_LIGHT

    # Guardar en buffer en memoria
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ------------------------------------------------------------------------------
# Interfaz Streamlit
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Generador PPTX IA Abastecimiento",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Generador de Presentación PPTX")
st.write(
    "Genera y descarga la presentación editable sobre "
    "**Transformación Digital y Agentes de IA en Abastecimiento**."
)

st.divider()

if st.button("🔨 Generar Presentación", type="primary"):
    with st.spinner("Construyendo la presentación..."):
        ppt_buffer = generar_presentacion_pptx()
    
    st.success("¡Presentación generada exitosamente!")
    
    st.download_button(
        label="📥 Descargar Abastecimiento_IA_Presentacion.pptx",
        data=ppt_buffer.getvalue(),
        file_name="Abastecimiento_IA_Presentacion.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
