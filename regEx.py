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

