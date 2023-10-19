import unicodedata

def process_strings(input_file, output_file):
    # Read the input file
    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # Process the strings
    processed_lines = []
    for line in lines:
        # Remove non-character symbols at the beginning
        processed_line = line.strip()
        for char in line:
            if unicodedata.category(char)[0] != 'L':
                processed_line = processed_line.lstrip(char)
            else:
                break
        
        # Make the first character lowercase
        processed_line = processed_line[0].lower() + processed_line[1:-1]+'\n'
        
        # Add the processed line to the list
        processed_lines.append(processed_line)

    # Write the modified list to the output file
    with open(output_file, 'w', encoding='utf-8',) as file:
        file.writelines(processed_lines)

# Example usage
input_file = 'input.txt'
output_file = 'output.txt'
process_strings(input_file, output_file)
