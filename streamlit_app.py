from io import BytesIO

import docx
import openpyxl
import pandas as pd
import streamlit as st
from google.cloud import translate_v2 as translate
from google.oauth2 import service_account

# Google Cloud Translation API setup
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"]
)
translate_client = translate.Client(credentials=credentials)

# Function to translate text


def translate_text(text, target_language, model="Google"):
    if model == "Google":
        result = translate_client.translate(
            text, target_language=target_language
        )
        return result["translatedText"]
    else:
        return "Unsupported translation model."

# Function to handle file upload and translation


def handle_file_upload(file, target_language, model="Google"):
    if file.type == "text/plain":
        content = file.read().decode("utf-8")
        translated_content = translate_text(content, target_language, model)
        return translated_content
    elif file.type == "text/csv":
        df = pd.read_csv(file)
        translated_df = df.applymap(
            lambda x: translate_text(x, target_language, model)
        )
        return translated_df.to_csv(index=False)
    elif file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        df = pd.read_excel(file)
        translated_df = df.applymap(
            lambda x: translate_text(x, target_language, model)
        )
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            translated_df.to_excel(writer, index=False, sheet_name="Sheet1")
            writer.save()
        processed_data = output.getvalue()
        return processed_data
    elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = docx.Document(file)

        # Translate paragraphs
        for para in doc.paragraphs:
            para.text = translate_text(para.text, target_language, model)

        # Translate tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = translate_text(
                        cell.text, target_language, model)

        output = BytesIO()
        doc.save(output)
        processed_data = output.getvalue()
        return processed_data
    else:
        return "Unsupported file type."


# Streamlit app UI
st.title("Document Translator")

uploaded_file = st.file_uploader(
    "Upload a document", type=["txt", "csv", "xlsx", "docx"]
)
languages = {
    "English": "en",
    "Spanish": "es",
    "French": "fr",
    "German": "de",
    "Chinese (Simplified)": "zh-CN",
    "Japanese": "ja",
    "Korean": "ko",
    "Portuguese": "pt",
    "Russian": "ru",
    "Arabic": "ar",
}

models = ["Google", "DeepL"]  # Add more models here as needed

language = st.selectbox("Select target language", list(languages.keys()))
model = st.selectbox("Select translation model", models)

if st.button("Translate"):
    if uploaded_file is not None:
        target_language = languages[language]
        translation = handle_file_upload(uploaded_file, target_language, model)
        if isinstance(translation, str):
            st.text_area("Translated Text", translation, height=300)
        else:
            st.download_button(
                label="Download Translated File",
                data=translation,
                file_name=f"translated_{uploaded_file.name}",
            )
    else:
        st.error("Please upload a file first.")
