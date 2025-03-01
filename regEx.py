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



######Getting Ch num and type for every writing style


# Function to extract chambre number and type
# def extract_info(text):
#     # Regex patterns
#     # chambre_pattern = re.compile(r'(\d+)(?:\/\d+)?')  # Matches any number of digits chamber number, optionally followed by / and more digits
#     # type_pattern = re.compile(r'([A-Za-z]\d[A-Za-z]?)')  # Matches chamber type (any letter, digit, letter)

#     combined_pattern = re.compile(r'(?:CH_FT/?)?(?:([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d)[\s-]+(\d+)(?:\/\d+)?|(\d+)(?:\/\d+)?(?:\s+([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d))?)')
#     match = combined_pattern.search(text)
#     # chambre_match = chambre_pattern.search(text)
#     # type_match = type_pattern.search(text)
    
#     chambre_number = match.group(2) if match else ""
#     chambre_type = match.group(1) if match else ""
    
#     return chambre_number, chambre_type



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






def extract_info(text):
    #pattern = re.compile(r'(?:CH_FT/?)?(?:([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d)(?:[\s-]+(\d+)(?:\/\d+)?)?|(\d+)(?:\/\d+)?(?:\s+([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d))?)')
    #pattern = re.compile(r'(?:CH_FT/?|CH/)?(?:([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d)(?:[\s-]+(\d+)(?:\/\d+)?)?|(\d+)(?:\/\d+)?(?:\s+([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d))?)')
    #pattern = re.compile(r'(?:CH_FT/?|CH/)?(?:([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d)(?:[\s-]+(?:FT/)?(\d+)(?:\/\d+)?)?|(\d+)(?:\/\d+)?(?:\s+([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d))?)')
    pattern = re.compile(r'(?:([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d)(?:[\s-]+(?:CH_FT/|FT/|CH/)?(\d+)(?:\/\d+)?)?|(?:CH_FT/|FT/|CH/)?(\d+)(?:\/\d+)?(?:\s+([A-Za-z]\d[A-Za-z]?|[A-Za-z]\d|[A-Za-z]{2}\d))?)')

    match = pattern.search(text)
    if match:
        # If type is before number (groups 1 and 2)
        if match.group(1):
            return match.group(2) or "", match.group(1) or ""
        # If type is after number (groups 3 and 4)
        elif match.group(3):
            return match.group(3) or "", match.group(4) or ""
    return "", ""  # Default to empty strings if no match

# # Test cases
# tests = [
#     "D7 35/48194",      # Should give ("35", "D7")
#     "L3T 34/48194",     # Should give ("34", "L3T")
#     "CH_FT 148 L3C",    # Should give ("148", "L3C")
#     "L2S-326/54698",    # Should give ("326", "L2S")
#     "326/54698",        # Should give ("326", None)
#     "CH_FT/",           # Should give (None, None)
#     "L2 CH_FT/",        # Should give (None, "L2")
#     "CH/00326  DR3",    # Should give ("00326", "DR3")
#     "L1T FT/206",       # Should give ("206", "L1T")
#     "L2C CH_FT/476"     # Should give ("476", "L2C")
# ]

# for test in tests:
#     chamber_num, type_code = extract_info(test)
#     print(f"Text: {test}")
#     print(f"Chamber: {chamber_num}, Type: {type_code}\n")
