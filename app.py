import os
import glob
import tempfile
import io
from contextlib import redirect_stdout, redirect_stderr
import streamlit as st

# Configuración de página
st.set_page_config(
    page_title="Ejecutor Python a PPTX",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Convertidor de Código Python a PPTX")
st.write(
    "Pega cualquier código Python que utilice `python-pptx`. "
    "La aplicación lo ejecutará en un entorno aislado, detectará la presentación generada y te permitirá descargarla."
)

# Código por defecto (tu script de ejemplo)
DEFAULT_CODE = '''import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

DARK_BG = RGBColor(15, 7, 10)
LIGHT_BG = RGBColor(255, 248, 249)
WHITE = RGBColor(255, 255, 255)
DARK_TEXT = RGBColor(17, 24, 39)
SUBTITLE_TEXT = RGBColor(100, 116, 139)
RED_PRIMARY = RGBColor(225, 29, 72)
RED_CRIMSON = RGBColor(244, 63, 94)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_slide_layout = prs.slide_layouts[6]

slide = prs.slides.add_slide(blank_slide_layout)
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = DARK_BG

tx_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.33), Inches(2.0))
tf = tx_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
p.text = "¡Presentación Generada desde Código!"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = WHITE

# Guardar la presentación (el nombre se detectará automáticamente)
output_filename = "mi_presentacion.pptx"
prs.save(output_filename)
print(f"Presentación guardada correctamente como: {output_filename}")
'''

col_editor, col_preview = st.columns([3, 2])

with col_editor:
    st.subheader("1. Pega tu código Python")
    user_code = st.text_area(
        "Código Python (.py):",
        value=DEFAULT_CODE,
        height=550,
        help="Asegúrate de incluir prs.save('nombre.pptx') al final de tu script."
    )
    
    run_button = st.button("🚀 Ejecutar Código y Generar PPTX", type="primary", use_container_width=True)

with col_preview:
    st.subheader("2. Resultado y Descarga")
    
    if run_button:
        if not user_code.strip():
            st.warning("Por favor, ingresa un código de Python válido.")
        else:
            with st.spinner("Ejecutando script..."):
                stdout_capture = io.StringIO()
                stderr_capture = io.StringIO()
                
                # Crear carpeta temporal aislada
                with tempfile.TemporaryDirectory() as temp_dir:
                    original_cwd = os.getcwd()
                    try:
                        # Cambiar al directorio temporal para que prs.save() guarde ahí
                        os.chdir(temp_dir)
                        
                        exec_globals = {}
                        
                        # Capturar salidas de consola (print/errores)
                        with redirect_stdout(stdout_capture), redirect_stderr(stderr_capture):
                            exec(user_code, exec_globals)
                        
                        # Buscar archivos .pptx creados durante la ejecución
                        pptx_files = glob.glob("*.pptx")
                        
                        if pptx_files:
                            filename = pptx_files[0]
                            with open(filename, "rb") as f:
                                file_bytes = f.read()
                            
                            st.success(f"✅ ¡Archivo `{filename}` generado exitosamente!")
                            
                            st.download_button(
                                label=f"📥 Descargar `{filename}`",
                                data=file_bytes,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ El script se ejecutó sin errores pero no generó ningún archivo `.pptx`. Revisa que incluyas `prs.save('nombre.pptx')`.")
                            
                    except Exception as e:
                        st.error(f"⚠️ Error en la ejecución del código:\n\n`{e}`")
                    finally:
                        # Restaurar directorio original
                        os.chdir(original_cwd)
                
                # Mostrar consola de salida (prints)
                logs = stdout_capture.getvalue()
                errs = stderr_capture.getvalue()
                
                if logs or errs:
                    with st.expander("📄 Ver logs de la consola (stdout/stderr)"):
                        if logs:
                            st.text("Salida estándar:")
                            st.code(logs)
                        if errs:
                            st.text("Errores / Advertencias:")
                            st.code(errs)
    else:
        st.info("Pega tu código a la izquierda y presiona el botón **Ejecutar Código**.")
