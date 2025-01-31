import streamlit as st



def upload_files(type,multiple=False):
    uploaded_files = st.file_uploader(
                f"test", type=[type],label_visibility="hidden", 
                accept_multiple_files=multiple,
            )
    
    return uploaded_files