import tempfile
from pathlib import Path

import streamlit as st


# Pasta onde estão main.py e template.yml
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.yml"


st.set_page_config(
    page_title="Conversor IES",
    page_icon="📄",
    layout="centered",
)


st.title("Conversor IES")
st.write("Selecione um ficheiro PDF para gerar o respetivo Excel.")


# Confirmar que o ficheiro de configuração existe
if not TEMPLATE_PATH.is_file():
    st.error(
        "O ficheiro template.yml não foi encontrado no servidor. "
        "Confirme que está na mesma pasta que o main.py."
    )
    st.stop()


# Estado utilizado para guardar o Excel após a conversão
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None

if "excel_name" not in st.session_state:
    st.session_state.excel_name = None

if "processed_file" not in st.session_state:
    st.session_state.processed_file = None


uploaded_file = st.file_uploader(
    label="PDF de entrada",
    type=["pdf"],
    accept_multiple_files=False,
)


if uploaded_file is not None:
    # Remove caminhos que eventualmente venham no nome do ficheiro
    original_name = Path(uploaded_file.name).name
    output_name = f"{Path(original_name).stem}.xlsx"

    # Identificador para detetar quando o utilizador troca de PDF
    uploaded_identifier = (
        original_name,
        uploaded_file.size,
    )

    if st.session_state.processed_file != uploaded_identifier:
        st.session_state.excel_data = None
        st.session_state.excel_name = None

    if st.button(
        "Converter para Excel",
        type="primary",
        use_container_width=False,
    ):
        st.session_state.excel_data = None
        st.session_state.excel_name = None

        try:
            with st.spinner("A preparar e converter o documento..."):
                import pikepdf
                from pdf_para_excel_template import main as converter_pdf_excel

                with tempfile.TemporaryDirectory() as temporary_directory:
                    temporary_path = Path(temporary_directory)

                    # Usamos nomes internos simples para evitar problemas
                    # com espaços, acentos ou outros caracteres.
                    source_pdf_path = temporary_path / "entrada.pdf"
                    prepared_pdf_path = temporary_path / "preparado.pdf"
                    excel_path = temporary_path / "resultado.xlsx"

                    source_pdf_path.write_bytes(uploaded_file.getvalue())

                    try:
                        # Abre e volta a guardar o PDF.
                        # Isto cria uma cópia temporária sem as restrições
                        # que bloqueiam a extração de texto pelo Camelot.
                        with pikepdf.open(
                            source_pdf_path,
                            password="",
                        ) as pdf:
                            pdf.save(prepared_pdf_path)

                    except pikepdf.PasswordError as error:
                        raise RuntimeError(
                            "O PDF exige uma palavra-passe para ser processado."
                        ) from error

                    except pikepdf.PdfError as error:
                        raise RuntimeError(
                            "O ficheiro não parece ser um PDF válido ou está danificado."
                        ) from error

                    # Executar a lógica original de conversão
                    converter_pdf_excel(
                        str(prepared_pdf_path),
                        str(TEMPLATE_PATH),
                        str(excel_path),
                        "overwrite",
                    )

                    if not excel_path.is_file():
                        raise RuntimeError(
                            "A conversão terminou, mas o ficheiro Excel não foi criado."
                        )

                    st.session_state.excel_data = excel_path.read_bytes()
                    st.session_state.excel_name = output_name
                    st.session_state.processed_file = uploaded_identifier

            st.success("Conversão concluída com sucesso.")

        except Exception as error:
            st.error(f"Não foi possível converter o PDF: {error}")


if st.session_state.excel_data is not None:
    st.download_button(
        label="Descarregar Excel",
        data=st.session_state.excel_data,
        file_name=st.session_state.excel_name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=False,
    )
