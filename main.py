import tempfile
from pathlib import Path

import streamlit as st

from pdf_para_excel_template import main as converter_pdf_excel


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.yml"


st.set_page_config(
    page_title="IES — PDF para Excel",
    page_icon="📄",
    layout="centered",
)

st.title("Conversor IES")
st.write("Selecione um ficheiro PDF para gerar o respetivo Excel.")

if not TEMPLATE_PATH.is_file():
    st.error("O ficheiro template.yml não foi encontrado no servidor.")
    st.stop()


uploaded_file = st.file_uploader(
    "PDF de entrada",
    type=["pdf"],
    accept_multiple_files=False,
)

if "excel_data" not in st.session_state:
    st.session_state.excel_data = None
    st.session_state.excel_name = None


if uploaded_file is not None:
    pdf_name = Path(uploaded_file.name).name
    output_name = f"{Path(pdf_name).stem}.xlsx"

    if st.button("Converter para Excel", type="primary"):
        st.session_state.excel_data = None
        st.session_state.excel_name = None

        try:
            with st.spinner("A converter o documento..."):
                with tempfile.TemporaryDirectory() as temp_directory:
                    temp_path = Path(temp_directory)

                    pdf_path = temp_path / pdf_name
                    excel_path = temp_path / output_name

                    pdf_path.write_bytes(uploaded_file.getvalue())

                    converter_pdf_excel(
                        str(pdf_path),
                        str(TEMPLATE_PATH),
                        str(excel_path),
                        "overwrite",
                    )

                    st.session_state.excel_data = excel_path.read_bytes()
                    st.session_state.excel_name = output_name

            st.success("Conversão concluída.")

        except Exception as error:
            st.error(f"Não foi possível converter o PDF: {error}")


if st.session_state.excel_data is not None:
    st.download_button(
        label="Descarregar Excel",
        data=st.session_state.excel_data,
        file_name=st.session_state.excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
