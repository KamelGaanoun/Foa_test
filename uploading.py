import streamlit as st



def upload_files(type,multiple=False):
    uploaded_files = st.file_uploader(
                f"Chargez un fichier Excel puis cliquez sur Traiter", type=[type], 
                accept_multiple_files=multiple,
            )
    
    return uploaded_files