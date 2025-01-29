
import openpyxl as px
from io import BytesIO

def load_excel_file(uploaded_files,read_only):
    # Use data_only=True to read formulas as values
    wb=px.load_workbook(BytesIO(uploaded_files.read()), data_only=True,read_only=read_only,keep_vba=False, keep_links=False)
    return wb