import re



def get_planche_number(text):
    # Regular expression to match the number pattern
    pattern = r"-(\d+)-"

    # Search for the first match
    match = re.search(pattern, text)
    if match:
        number = match.group(1)
        return number
    else:
        return None


# Function to modify the text based on the number after "_"
def process_text(text):
    match = re.match(r"(\d+)_(\d+)\.jpg", text)  # Match the pattern 723784_1.jpg
    if match:
        base, num = match.groups()
        if num == "1":
            return f"{base}_3.jpg"
        elif num == "2":
            return f"{base}_4.jpg"
    return text  # Return original if no match
