# app.py
import streamlit as st
import os
from io import BytesIO
import msg_converter_core # Import our conversion logic

st.set_page_config(page_title="MSG to EML Converter", layout="wide")

st.title("✉️ MSG to EML Converter (with Nested Attachments)")
st.markdown("""
Upload a `.msg` file. This tool will convert it into a single `.eml` file,
with any nested `.msg` attachments also converted and embedded as `message/rfc822` parts
(viewable as attached EMLs in most mail clients).
""")

uploaded_file = st.file_uploader("Choose a .MSG file", type=["msg"])

if uploaded_file is not None:
    file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
    st.write("---")
    st.subheader("Uploaded File Details:")
    st.json(file_details)

    # Use BytesIO to handle the uploaded file bytes as a file-like object
    msg_bytes_io = BytesIO(uploaded_file.getvalue())
    original_filename_stem = os.path.splitext(uploaded_file.name)[0]

    st.write("---")
    st.subheader("Conversion Process:")
    
    # Placeholder for logs
    log_area = st.empty()
    logs = []
    def streamlit_log_callback(message):
        logs.append(message)
        log_area.code("\n".join(logs)) # Display logs in a code block

    # Perform conversion
    if st.button(f"Convert '{uploaded_file.name}' to EML"):
        logs = [] # Clear previous logs
        streamlit_log_callback("Starting conversion...")
        
        with st.spinner("Processing your .msg file... This might take a moment for complex files."):
            eml_bytes, suggested_download_name = msg_converter_core.convert_msg_to_single_eml(
                msg_bytes_io, 
                original_filename_stem,
                log_callback=streamlit_log_callback
            )

        if eml_bytes and suggested_download_name:
            streamlit_log_callback(f"Conversion successful! ✅")
            st.balloons()
            st.subheader("Download Your EML File")
            st.download_button(
                label=f"Download {suggested_download_name}",
                data=eml_bytes,
                file_name=suggested_download_name,
                mime="message/rfc822" # Correct MIME type for .eml files
            )
        else:
            streamlit_log_callback("Conversion failed. 😔 Please check the logs above.")
            st.error("Conversion failed. See logs for details.")
else:
    st.info("Upload a .msg file to begin.")

st.write("---")
st.markdown("Developed with Streamlit and extract-msg.")