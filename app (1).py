
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
import io
import zipfile
import os
# from google.colab import userdata # Import userdata - REMOVIDO para compatibilidade com Streamlit Cloud

# Caminho do logótipo
LOGO_PATH = "esposack-logo.png" # Assumes the logo is in the same directory or accessible path

# Função para gerar certificados
def gerar_certificado(cliente, guia, data, artigo, descricao, lotes, pdf_buffer, logo_path=LOGO_PATH):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Body', fontName='Helvetica', fontSize=10.5, leading=14))
    styles.add(ParagraphStyle(name='BodyBold', fontName='Helvetica-Bold', fontSize=10.5, leading=14))
    styles.add(ParagraphStyle(name='TitlePT', fontName='Helvetica-Bold', fontSize=16, leading=20, alignment=1, spaceAfter=12))
    styles.add(ParagraphStyle(name='Justify', fontName='Helvetica', fontSize=10.5, leading=15, alignment=4))
    styles.add(ParagraphStyle(name='Right', fontName='Helvetica', fontSize=10.5, alignment=2))

    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []

    # Cabeçalho
    try:
        # Verificar se o ficheiro do logótipo existe antes de tentar carregá-lo
        if os.path.exists(logo_path):
            logo_obj = Image(logo_path, width=35*mm, height=20*mm)
            address_para = Paragraph(
                "ESPOSACK – EMBALAGENS Lda.<br/>"
                "Zona Industrial do Bouro, Pav. nº6<br/>"
                "4740-010 Gandra - Esposende<br/>"
                "Telefone: 253 962 064<br/>"
                "www.esposack.pt",
                styles['Body']
            )
            header_tbl = Table([[address_para, logo_obj]], colWidths=[None, 45*mm])
            header_tbl.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
            ]))
            story.append(header_tbl)
            story.append(Spacer(1, 12))
        else:
            # Adicionar apenas as informações da empresa se o logótipo não for encontrado
            address_para = Paragraph(
                "ESPOSACK – EMBALAGENS Lda.<br/>"
                "Zona Industrial do Bouro, Pav. nº6<br/>"
                "4740-010 Gandra - Esposende<br/>"
                "Telefone: 253 962 064<br/>"
                "www.esposack.pt",
                styles['Body']
            )
            story.append(address_para)
            story.append(Spacer(1, 12))
            st.warning(f"Logótipo não encontrado em {logo_path}. Gerando certificado sem logótipo.") # Usar st.warning no Streamlit

    except Exception as e:
        st.error(f"Erro ao carregar ou adicionar logótipo: {e}") # Usar st.error no Streamlit
        story.append(Paragraph("Certificado de Conformidade", styles['TitlePT']))
        story.append(Spacer(1, 12))


    # Conteúdo principal
    story.append(Paragraph("Certificado de Conformidade", styles['TitlePT']))
    story.append(Spacer(1, 6))

    line1_tbl = Table([
        [Paragraph(f"Cliente: <b>{cliente}</b>", styles['Body']),
         Paragraph(f"Guia de Remessa: <b>{guia}</b>", styles['Body'])]
    ], colWidths=[270, None])
    line1_tbl.setStyle(TableStyle([('ALIGN', (1,0), (1,0), 'RIGHT')]))
    story.append(line1_tbl)
    story.append(Spacer(1, 6))

    story.append(Paragraph(f"Data: <b>{data}</b>", styles['Body']))
    story.append(Paragraph(f"Artigo: <b>{artigo}</b> &nbsp;&nbsp;&nbsp; Descrição do Artigo: <b>{descricao}</b>", styles['Body']))
    story.append(Spacer(1, 10))

    # Tabela de lotes
    table_data = [[Paragraph("<b>Lote</b>", styles['BodyBold']),
                   Paragraph("<b>Quantidade</b>", styles['BodyBold']),
                   Paragraph("<b>Unidade</b>", styles['BodyBold'])]] + lotes

    tbl = Table(table_data, colWidths=[100, 100, 100])
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.6, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 18))

    texto = (
        "Declara-se que os lotes acima identificados foram produzidos conforme as "
        "especificações técnicas acordadas, de acordo com o Regulamento (CE) Nº. 1935/2004 "
        "de 27 de Outubro de 2004, Diretivas 2007/42/CE da Comissão de 20 de Junho de 2007, "
        "transposta para legislação nacional pelo Decreto-Lei nº 194/2007, de 14 de Maio e de "
        "acordo com os requisitos FDA Parágrafo 176.170/21 CFR. Os lotes acima identificados "
        "passaram pelo controlo de qualidade da Esposack."
    )
    story.append(Paragraph(texto, styles['Justify']))
    story.append(Spacer(1, 36))

    story.append(Paragraph("Alexandra Garrido", styles['Right']))
    story.append(Paragraph("(Dept. Qualidade)", styles['Right']))

    doc.build(story)


# Set the title of the web application
st.title("Gerador de Certificados Esposack")

# Add a file uploader widget that accepts only .xlsx files
uploaded_file = st.file_uploader(
    "Carregue o ficheiro Excel (formato .xlsx)", type=["xlsx"]
)

df = None # Initialize df outside the if block

if uploaded_file is not None:
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)
        st.success("Ficheiro carregado com sucesso!")
        st.write("Pré-visualização dos dados:")
        st.dataframe(df.head()) # Use st.dataframe() to display the DataFrame

        # Process data and generate certificates if DataFrame is not empty
        if df is not None and not df.empty:
            # Group by Guia to create separate zip files for each
            if "Nº Guia" in df.columns:
                grouped_by_guia = df.groupby("Nº Guia")

                for guia, guia_group in grouped_by_guia:
                    pdf_buffers = {}
                    # Group by relevant columns within each Guia group
                    # Ensure all grouping columns exist before grouping
                    grouping_cols = ["Cliente", "Nº Guia", "Data", "Artigo", "Descrição do Artigo"]
                    if all(col in guia_group.columns for col in grouping_cols):
                         grouped_deliveries = guia_group.groupby(grouping_cols)

                         for (cliente, current_guia, data, artigo, descricao), grupo in grouped_deliveries:
                             lotes = grupo[["Lote", "Quantidade", "Unidade"]].values.tolist()
                             pdf_buffer = io.BytesIO()
                             # Ensure data is in string format, especially for the date
                             data_str = data.strftime('%d/%m/%Y') if isinstance(data, pd.Timestamp) else str(data)
                             gerar_certificado(cliente, current_guia, data_str, artigo, descricao, lotes, pdf_buffer, LOGO_PATH)
                             pdf_buffer.seek(0) # Rewind the buffer
                             # Use a filename that makes sense for a certificate per article
                             filename = f"Certificado_Guia_{current_guia}_Artigo_{artigo}.pdf"
                             pdf_buffers[filename] = pdf_buffer

                         st.success(f"Certificados gerados com sucesso para a Guia {guia}!")

                         # Create a zip file in memory for this Guia
                         zip_buffer = io.BytesIO()
                         zip_filename = f"Certificados_esposack_Guia_{guia}.zip"
                         with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                             for filename, pdf_buffer in pdf_buffers.items():
                                 zip_file.writestr(filename, pdf_buffer.getvalue())

                         zip_buffer.seek(0) # Rewind the buffer

                         # Provide a download button for this Guia's zip file
                         st.download_button(
                             label=f"Descarregar Certificados Guia {guia} (ZIP)",
                             data=zip_buffer,
                             file_name=zip_filename,
                             mime="application/zip"
                         )
                    else:
                        missing_cols = [col for col in grouping_cols if col not in guia_group.columns]
                        st.warning(f"Grupo da Guia {guia} ignorado: Faltam as colunas necessárias para agrupar: {', '.join(missing_cols)}")

                st.write("Processamento de todas as guias concluído.")

            else:
                 st.error("A coluna 'Nº Guia' não foi encontrada no ficheiro Excel. Certifique-se de que o nome da coluna está correto.")


        else:
            st.info("O ficheiro Excel está vazio ou não contém os dados esperados.")

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler ou processar o ficheiro Excel: {e}")

else:
    st.info("Por favor, carregue o ficheiro Excel com os dados da entrega para gerar os certificados.")

