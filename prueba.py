import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile
import io

# --- Funciones auxiliares ---

def reemplazar_texto(doc, reemplazos):
    for p in doc.paragraphs:
        for clave, valor in reemplazos.items():
            if clave in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if clave in inline[i].text:
                        inline[i].text = inline[i].text.replace(clave, valor)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for clave, valor in reemplazos.items():
                    if clave in cell.text:
                        cell.text = cell.text.replace(clave, valor)

def eliminar_parrafos(doc, frases_a_eliminar):
    for para in doc.paragraphs:
        for frase in frases_a_eliminar:
            if frase in para.text:
                p = para._element
                p.getparent().remove(p)
                break

def generar_documentos(excel_file, word_template):
    df = pd.read_excel(excel_file)

    output_folder = "documentos_generados"
    os.makedirs(output_folder, exist_ok=True)

    generated_files = []

    for idx, row in df.iterrows():
        doc = Document(word_template)

        # Diccionario de reemplazos
        reemplazos = {
            "<<NUMERO_PROTOCOLO>>": str(row["Numero de protocolo"]),
            "<<TITULO_ESTUDIO>>": str(row["Titulo del Estudio"]),
            "<<PATROCINADOR>>": str(row["Patrocinador"]),
            "<<INVESTIGADOR>>": str(row["Investigador"]),
            "<<INSTITUCION>>": str(row["Institucion"]),
            "<<DIRECCION>>": str(row["Direccion"]),
            "<<CARGO_INVESTIGADOR>>": str(row["Cargo del Investigador en la Institucion"]),
        }

        reemplazar_texto(doc, reemplazos)

        # Si es Córdoba, borrar párrafos
        if "cordoba" in str(row["Direccion"]).lower():
            frases_a_eliminar = [
                "El medico del estudio discutira con Usted que metodo anticonceptivo",
                "requerido para centros de la provincia de buenos aires"
            ]
            eliminar_parrafos(doc, frases_a_eliminar)

        # Nombre de archivo
        protocolo = str(row["Numero de protocolo"]).replace(" ", "_")
        investigador = str(row["Investigador"]).replace(" ", "_")
        output_filename = f"consentimiento_{protocolo}_{investigador}.docx"
        output_path = os.path.join(output_folder, output_filename)

        doc.save(output_path)
        generated_files.append(output_path)

    # Crear ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for file in generated_files:
            zipf.write(file, os.path.basename(file))

    zip_buffer.seek(0)
    return zip_buffer


# --- Interfaz con Streamlit ---

st.title("📄 Generador Automático de Consentimientos Informados")

st.write("Subí el archivo **Excel** con los datos de los centros y el archivo **Word modelo** con los placeholders.")

excel_file = st.file_uploader("📑 Subir Excel (centros.xlsx)", type=["xlsx"])
word_template = st.file_uploader("📝 Subir modelo (modelo.docx)", type=["docx"])

if excel_file and word_template:
    if st.button("⚙️ Generar documentos"):
        with st.spinner("Generando documentos..."):
            zip_file = generar_documentos(excel_file, word_template)

        st.success("✅ Documentos generados correctamente")
        st.download_button(
            label="⬇️ Descargar ZIP con documentos",
            data=zip_file,
            file_name="documentos_generados.zip",
            mime="application/zip"
        )
