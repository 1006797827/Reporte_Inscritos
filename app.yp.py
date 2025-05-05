import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import io
import base64

# Configuración de la página
st.set_page_config(
    page_title="Generador de Reportes",
    page_icon="📊",
    layout="wide"
)

# Título y descripción
st.title("Generador de Reportes de Inscripciones")
st.markdown("""
Esta aplicación genera un informe PDF a partir de datos de inscripción a cursos.
Sube un archivo Excel con las columnas: "Nombre y apellidos completos", "Hora de inicio", 
"Curso de interés", "Correo de contacto" y "Número de contacto".
""")

# Función para generar el informe PDF
def generar_informe_pdf(df):
    # Obtener la cantidad de personas inscritas (valores únicos)
    num_inscritos = len(df["Nombre y apellidos completos"].unique())
    
    # Obtener las fechas máxima y mínima de inscripción
    fecha_minima = df["Hora de inicio"].min()
    fecha_maxima = df["Hora de inicio"].max()
    
    # Obtener la fecha y hora actuales para el encabezado del informe
    fecha_hora_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Crear el buffer para el PDF
    buffer = io.BytesIO()
    
    with PdfPages(buffer) as pdf:
        # --- Primera página ---
        fig, ax = plt.subplots(figsize=(8.27, 11.69))  # Tamaño carta
        ax.axis('off')  # Desactivar los ejes
        
        # Agregar un marco
        ax.patch.set_edgecolor('black')  
        ax.patch.set_linewidth(2)
        
        # Agregar título y contenido
        ax.text(0.5, 0.95, "INFORME DE PERSONAS INSCRITAS", fontsize=16, fontweight='bold', ha='center', va='top')
        ax.text(0.5, 0.90, "___________________________________________", fontsize=12, ha='center', va='top')
        ax.text(0.05, 0.80, f"Número de inscritos: {num_inscritos}", fontsize=12, ha='left', va='top')
        ax.text(0.05, 0.75, f"Fecha de inicio de inscripciones: {fecha_minima}", fontsize=12, ha='left', va='top')
        ax.text(0.05, 0.70, f"Fecha de fin de inscripciones: {fecha_maxima}", fontsize=12, ha='left', va='top')
        ax.text(0.05, 0.60, f"Fecha y hora de elaboración: {fecha_hora_actual}", fontsize=10, ha='left', va='top')
        ax.text(0.05, 0.55, "Elaborado por: Laura Valentina Jimenez Benavides", fontsize=10, ha='left', va='top')
        
        plt.subplots_adjust(top=0.95, bottom=0.05)
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)
        
        # --- Segunda página: Gráfica de barras ---
        inscripciones_por_curso = df.groupby("Curso de interés")["Nombre y apellidos completos"].count()
        
        fig, ax = plt.subplots(figsize=(8.27, 11.69))
        
        # Crear gráfico de barras con colores diferentes
        bars = inscripciones_por_curso.plot(
            kind="bar", 
            ax=ax, 
            color=plt.cm.get_cmap("tab20")(range(len(inscripciones_por_curso)))
        )
        
        # Configurar etiquetas y título
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        
        # Añadir valores en las barras
        for i, v in enumerate(inscripciones_por_curso):
            ax.text(i, v + 0.5, str(v), ha='center')
            
        ax.set_xlabel("Curso de interés")
        ax.set_ylabel("Número de inscritos")
        ax.set_title("Inscripciones por curso")
        
        # Ajustar márgenes para asegurar que todo quepa
        plt.subplots_adjust(bottom=0.3)
        
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)
        
        # --- Tercera página en adelante: Tablas por curso ---
        for curso in df["Curso de interés"].unique():
            df_curso = df[df["Curso de interés"] == curso]
            
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis('off')
            
            # Crear tabla
            tabla = ax.table(
                cellText=df_curso[["Nombre y apellidos completos", "Correo de contacto", "Número de contacto"]].values,
                colLabels=["Nombre y apellidos completos", "Correo de contacto", "Número de contacto"],
                loc='center', 
                cellLoc='center'
            )
            
            tabla.auto_set_font_size(False)
            tabla.set_fontsize(8)
            tabla.scale(1, 1.5)  # Ajustar escala para mejor visualización
            
            # Ajustar ancho de columnas
            tabla.auto_set_column_width([0, 1, 2])
            
            ax.set_title(f"Personas inscritas en {curso}", pad=20)
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)
    
    buffer.seek(0)
    return buffer

# Widget para cargar archivo
uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Cargar los datos
        df = pd.read_excel(uploaded_file)
        
        # Verificar que el archivo tenga las columnas requeridas
        required_columns = [
            "Nombre y apellidos completos", 
            "Hora de inicio", 
            "Curso de interés", 
            "Correo de contacto", 
            "Número de contacto"
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"El archivo no contiene las siguientes columnas requeridas: {', '.join(missing_columns)}")
        else:
            # Mostrar vista previa de los datos
            st.subheader("Vista previa de los datos")
            st.dataframe(df.head())
            
            # Mostrar estadísticas básicas
            st.subheader("Estadísticas de inscripción")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de inscritos", len(df["Nombre y apellidos completos"].unique()))
            with col2:
                st.metric("Número de cursos", len(df["Curso de interés"].unique()))
            with col3:
                st.metric("Periodo de inscripción", f"{df['Hora de inicio'].min().date()} - {df['Hora de inicio'].max().date()}")
            
            # Botón para generar informe
            if st.button("Generar Informe PDF"):
                with st.spinner("Generando informe..."):
                    pdf_buffer = generar_informe_pdf(df)
                    
                    # Codificar el PDF para descarga
                    b64_pdf = base64.b64encode(pdf_buffer.read()).decode('utf-8')
                    href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="Reporte_Personas_Inscritas.pdf" class="btn">⬇️ Descargar Reporte PDF</a>'
                    
                    st.success("¡Informe generado con éxito!")
                    st.markdown(href, unsafe_allow_html=True)
                    
                    # Vista previa del PDF
                    st.subheader("Vista previa del informe")
                    st.markdown(f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="500" type="application/pdf"></iframe>', unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")