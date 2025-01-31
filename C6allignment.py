from collections import defaultdict
from io import BytesIO
import zipfile



def extract_images(sheet):

    # Dictionary to store images grouped by row
    images_by_row = defaultdict(list)

    # Extract image positions and store them by row, sorting them by column
    for image in sheet._images:
        if hasattr(image.anchor, "_from"):  # Ensure valid anchor
            row = image.anchor._from.row
            col = image.anchor._from.col
            images_by_row[row].append((col, image))  # Store images by row

    # Sort images by column within each row
    for row in images_by_row:
        images_by_row[row] = sorted(images_by_row[row], key=lambda x: x[0])


    images_by_row = dict(sorted(images_by_row.items()))

    return images_by_row


def extract_texts(sheet):

    # Dictionary to store text cells grouped by row
    texts_by_row = defaultdict(list)

    # Get max column and row count
    max_row = sheet.max_row
    max_col = sheet.max_column

    # Extract text cells and group them by row
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value:  # Check if cell is not empty
                texts_by_row[row].append((col, cell.value))

    # Sort text entries by column within each row
    for row in texts_by_row:
        texts_by_row[row] = sorted(texts_by_row[row], key=lambda x: x[0])

    # Remove the first two rows from texts_by_row
    texts_by_row = {row: texts for row, texts in texts_by_row.items() if row > sorted(texts_by_row.keys())[1]}

    texts_by_row = {row - 2: texts for row, texts in texts_by_row.items()}

    return texts_by_row



def allign_images(images_by_row,sheet):
    # Create a list of keys to remove from images_by_row
    keys_to_remove = []

    texts_by_row=extract_texts(sheet)

    # Iterate through images_by_row
    for key in list(images_by_row.keys()):
        if key not in texts_by_row:
            # Find the next available key in texts_by_row
            next_key = key + 1
            while next_key not in texts_by_row:
                next_key += 1
            
            # Move the content to the next available key
            if next_key in images_by_row:
                # Prepend the new content to the existing content
                images_by_row[next_key] = f"{images_by_row[key]}, {images_by_row[next_key]}"
            else:
                # Assign the content to the next key
                images_by_row[next_key] = images_by_row[key]
            
            # Mark the key for removal
            keys_to_remove.append(key)

    # Remove the keys from images_by_row
    for key in keys_to_remove:
        del images_by_row[key]

    return images_by_row





# Function to flatten the dictionary and extract images
def flatten_images(images_dict):
    flattened_images = []
    for value in images_dict.values():
        if isinstance(value, str):
            # Handle string representation of lists
            try:
                # Convert the string to a list of tuples
                value = eval(value)
            except:
                continue
        if isinstance(value, list):
            # Extract the image objects from the list of tuples
            for _, image in value:
                flattened_images.append(image)
    return flattened_images




# Function to flatten the dictionary and extract .jpg text entries
def flatten_texts(texts_dict):
    flattened_texts = []
    for value in texts_dict.values():
        for _, text in value:
            if isinstance(text, str) and text.endswith('.jpg'):
                flattened_texts.append(text)
    return flattened_texts


# Function to save images into a ZIP file
def save_images_to_zip(images, titles):
    # Create a BytesIO buffer for the ZIP file
    zip_buffer = BytesIO()

    # Create a ZIP file
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for image, title in zip(images, titles):
            # Access the image data using the _data() method
            image_data = image._data()  # Call the _data method to get the bytes

            # Add the image to the ZIP file
            zip_file.writestr(title, image_data)

    # Return the ZIP file buffer
    zip_buffer.seek(0)
    return zip_buffer