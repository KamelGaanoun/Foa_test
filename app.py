
import streamlit as st
import  streamlit_toggle as tog
import openpyxl as px
from htmlTemplates import css,upload_style
from io import BytesIO
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import json
import shutil
import os
from reading import load_excel_file
from uploading import upload_files
from regEx import get_planche_number
from insert import foa_feeder



def main():

    st.set_page_config(page_title="FOA Builder", page_icon=":satellite_antenna:")
    
    # Inject the custom CSS to override Streamlit file uploader
    st.markdown(upload_style,unsafe_allow_html=True,)
    st.write(css, unsafe_allow_html=True)


    st.header("Créez votre FOA en un clic :satellite_antenna:")


    
    #Uploading the file
    uploaded_files=upload_files(type="xlsx",multiple=False)


    #Toggle to save images or not
    image_save=st.toggle("Voulez-vous sauvegarder les photos ?",value=False)


    if uploaded_files:
        #Gettin workbook from the session, otherwise st keeps refreshing and reload 
        #the wb each time it refreshes, and decreases user exp quality
        if "workbook" not in st.session_state:
            with st.spinner("Lecture et validation du fichier..."):
                wb = load_excel_file(uploaded_files,read_only=True)  # Load the Excel file
                st.session_state.workbook = wb  # Store the workbook in session state
                st.session_state.sheet_names = wb.sheetnames  # Store sheet names

        # Access stored workbook and sheet names
        wb = st.session_state.workbook
        sheet_names = st.session_state.sheet_names

         # Add an empty default choice (So the user is encouraged to correctly choose
         #the correct sheet name, otherwise maybe he will inadvertly let the first option
         #on the menu, which is not surely the good one)
        sheet_options = [""] + sheet_names  # Prepend an empty choice

        #Get nmero commande
        command_num=st.text_input("Merci de renseigner le numéro de commande")        
        
        if command_num:
            # Step 2: Prompt for the sheet name and provide a button to process
            sheet_name = st.selectbox("Veuillez sélectionner la feuille contenant les photos puis cliquez sur Traiter", sheet_options)
            process_button = st.button("Traiter")

            wb.close()  # Close the workbook to free up memory

            #What happens when "Traiter" is clicked
            if process_button:
                # Check if the user has selected a valid sheet name
                if sheet_name == "":
                    st.warning("Veuillez sélectionner une feuille avant de continuer.")
                else:
                    with st.spinner("Traitement en cours..."):
                        try:
                            wb = load_excel_file(uploaded_files,read_only=False)
                            #get the selected sheet
                            sheet = wb[sheet_name]

                            # Store PIL images in memory
                            pil_images_in_memory = []

                            # Create a directory with the same name as the uploaded file (without the extension)
                            base_filename = os.path.splitext(uploaded_files.name)[0]  # Remove .xlsx extension
                            save_directory = os.path.join("Images", base_filename)
                            outputs_directory = os.path.join("Outputs", base_filename)
                            for directory in [save_directory, outputs_directory]:
                                os.makedirs(directory, exist_ok=True)

                                
                            
                            # Extract images along with their positions (row and column)
                            # images_with_positions = []
                            # for image in sheet._images:
                            #     # Extract row and column from the image's anchor
                            #     row = image.anchor._from.row
                            #     col = image.anchor._from.col
                            #     images_with_positions.append((row, col, image))

                            # # Sort images by row (top-to-bottom) and column (left-to-right)
                            # images_with_positions.sort(key=lambda x: (x[0], x[1]))

                            images_with_positions = sorted(
                            ((image.anchor._from.row, image.anchor._from.col, image) for image in sheet._images),
                            key=lambda x: (x[0], x[1])
                            )
                                                
                            
                            
                            #Track col number (see just below)
                            col_number=-1

                            # Initialize database (where images metadata will be stored)
                            image_database = []
                            # Dictionary to store PIL images in memory with unique IDs
                            image_memory = {}

                            for idx, (row, col, image) in enumerate(images_with_positions):
                                
                                #Track col number, as some photos are in the same col so we have to increment
                                #it for the uniqueness of image names
                                if col_number==col:
                                    col_number+=1
                                else:
                                    col_number=col
                                # Extract the binary content of the image
                                img_bytes = image.ref.getvalue()  # Get the image binary data
                                
                                # Convert to a PIL image for display
                                pil_image = PILImage.open(BytesIO(img_bytes))
                                
                                
                                

                                # Retrieve textual information 
                                if row > 1:  # Ensure the row exists (skip if it's the first row)
                                    
                                    PB_data = [
                                        sheet.cell(row=row-3 , column=c).value
                                        for c in range(1, sheet.max_column + 1)
                                    ]
                                    Address_data = [
                                        sheet.cell(row=row , column=c).value
                                        for c in range(1, sheet.max_column + 1)
                                    ]

                                    
                                    #st.write(f"Textual Information (Row {row - 3}):", PB_data)
                                    #st.write(f"Textual Information (Row {row }):", Address_data)
                                
                                
                                # Save the image to the directory with a unique name based on row and column
                                
                                #Replace spaces(not convenient) and slashes(not allowed) with "_" as they are not allowed in file names
                                id_address=Address_data[3].replace(" ","_").replace("/","_").replace("\t","")
                                
                                img_title=f"image_{row}_{col_number}_{id_address}.png"
                                img_filename = os.path.join(save_directory, img_title)
                                
                                
                                # Store image in memory with its unique ID
                                image_memory[img_title] = pil_image

                                

                                if image_save:
                                    pil_image.save(img_filename)  # Save as PNG


                                ##Save Image info to DB

                                # Flatten PB_Data and Address_data into a single dictionary
                                flattened_data = {}
                                
                                #Get Planche number from the file title
                                planche=get_planche_number(uploaded_files.name)

                                # Add the image identifier and metadata into the list
                                flattened_data["id"] = img_title  # Add the image name as an identifier
                                flattened_data["planche"]= planche
                                
                                # Flatten PB_Data and Address_data into a single dictionary
                                flattened_data.update({
                                    key.replace(":", "").strip(): value.strip()
                                    for key, value in zip(PB_data[::2], PB_data[1::2])
                                })
                                flattened_data.update({
                                    key.replace(":", "").strip(): value.strip()
                                    for key, value in zip(Address_data[::2], Address_data[1::2])
                                })

                                image_database.append(flattened_data)

                                                                
                                # Display Image data
                                #st.json(flattened_data)
                            
                            
                            #Transform the db into a structure avoiding redundancy
                            # Group entries by their unique metadata values
                            grouped_db = {}
                            for entry in image_database:
                                # Create a unique key based on metadata values
                                metadata_key = (entry["planche"], entry["PB"], entry["Emplacement"], entry["Adresse"], entry["Site Support"])
                                
                                # If the group doesn't exist, create it
                                if metadata_key not in grouped_db:
                                    grouped_db[metadata_key] = []
                                
                                # Append the image ID to the group
                                grouped_db[metadata_key].append({"id": entry["id"]})

                            # Convert the grouped database into the new structured format
                            new_db = []
                            for metadata_key, images in grouped_db.items():
                                metadata = {
                                    "planche": metadata_key[0],
                                    "PB": metadata_key[1],
                                    "Emplacement": metadata_key[2],
                                    "Adresse": metadata_key[3],
                                    "Site Support": metadata_key[4]
                                }
                                new_db.append({
                                    "metadata": metadata,
                                    "images": images
                                })


                            # Save the database as a JSON file
                            database_path = os.path.join("data", "image_database.json")
                            with open(database_path, "w", encoding="utf-8") as json_file:
                                json.dump(new_db, json_file, ensure_ascii=False, indent=4)

                            if image_save:
                                st.toast(f"Les photos ont été récupérées et sauvegardés sur: {save_directory}")
                            else:
                                st.toast(f"Les photos ont été récupérées avec succès!" ,icon="✅")

                                # Display the image
                                #st.write(f"Image at Row: {row}, Column: {col}")
                                #st.image(pil_image)
                            
                            #foa_feeder(new_db,"00_C16.xlsx",image_memory,outputs_directory)#desktop
                            foa_feeder(new_db,"00_C16.xlsx",command_num,image_memory,base_filename)#web
                            
                            if "zip_ready" in st.session_state and st.session_state["zip_ready"]:
                                zip_file_path = st.session_state["zip_file_path"]

                                if os.path.exists(zip_file_path):
                                    with open(zip_file_path, "rb") as zipf:
                                        zip_data = zipf.read()  # Read the file before passing it to the button
                                        
                                        st.download_button(
                                            label="📥 Téléchargez vos fichiers",
                                            data=zip_data,  # Pass the file content, not the open file
                                            file_name=os.path.basename(zip_file_path),
                                            mime="application/zip",
                                            key="download_zip"  # Unique key to prevent disappearance
                                        )

                                    # Ensure the button remains visible after first display
                                    st.session_state["show_download"] = True

                                # if st.session_state.get("show_download", False):
                                #     st.success("✅ Vos fichiers sont prêts! Cliquez ci-dessus pour télécharger.")
                        except Exception as e:
                            st.error(f"An error occurred while processing the file: {e}")
                
        #else:
            #st.warning("Merci de charger un fichier Excel")
            
        




if __name__ == "__main__":
    main()



