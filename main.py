import xlrd
from sympy import Matrix
import hashlib
workbook = xlrd.open_workbook('ATTENDANCE_SHEET.xls')
sheet = workbook.sheet_by_name('DynamicReport')


def map_string_to_integer(input_string):
    # Generate a hash from the input string
    hash_value = int(hashlib.md5(input_string.encode()).hexdigest(), 16)

    # Map the hash value to the desired range (1 to 50)
    mapped_value = (hash_value % 50) + 1  # Ensure the result is between 1 and 50

    return mapped_value


def generate_unique_list(d1, d2, d3, x):
    # Concatenate the data and generate a hash
    hash_input = str(d1) + str(d2) + str(d3) + str(x)
    hash_value = int(hashlib.md5(hash_input.encode()).hexdigest(), 16)

    # Generate a list of 6 unique elements between 0 to 9 using the hash value
    unique_elements = []
    while len(unique_elements) < 6:
        num = hash_value % 10  # Get the last digit of the hash value (0-9)
        if num not in unique_elements:
            unique_elements.append(num)
        hash_value //= 10  # Move to the next digit

    return unique_elements

def map_string_to_ABC_DEF(input_string, x):
    hash_value = int(hashlib.md5(input_string.encode()).hexdigest(), 16)
    mapped_value = hash_value % 3
    if x==1:
        if mapped_value == 0:
            return 'A'
        elif mapped_value == 1:
            return 'B'
        else:
            return 'C'
    else:
        if mapped_value == 0:
            return 'D'
        elif mapped_value == 1:
            return 'E'
        else:
            return 'F'


def add_text_in_file(ID, Name):
    try:
        with open('origin.txt', 'r') as file:
            # Read the content of the file
            file_content = file.read()
        
        
        
        
        print(Name)
        # Replace the old text with the new text
        file_content = file_content.replace('@Name@', Name)
        file_content = file_content.replace('@ID@', str(ID))
        
        d1 = int((ID%1000)/100)
        d2 = int((ID%100)/10)
        d3 = (ID%10)
        
        for i in range(6):
            input_list = generate_unique_list(d1, d2, d3, i+1)
            set_format = "{" + ", ".join(str(item) for item in input_list) + "}"
            file_content = file_content.replace('@S' + str(i+1) + '@', set_format)
        
        for i in range(18):
            out_int = map_string_to_integer(str(ID) + str(i))
            file_content = file_content.replace('@g' + str(i+1) + '@', str(out_int))
        
        file_content = file_content.replace('@Row@', map_string_to_ABC_DEF(str(ID), 1))
        file_content = file_content.replace('@Col@', map_string_to_ABC_DEF(str(ID), 2))




        # Write the updated content back to the file
        with open('file.txt', 'a') as file:
            file.write(file_content)

        print('Text replacement successful.')
    except FileNotFoundError:
        print(f'The file "{file_path}" was not found.')
    except Exception as e:
        print(f'An error occurred: {str(e)}')




for i in range(44):
    row_number = 3+i
    ID = int(sheet.cell_value(row_number, 2))
    Name = sheet.cell_value(row_number, 3)
    add_text_in_file(ID, Name)


