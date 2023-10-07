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
        
        
        
        
        # Replace the old text with the new text
        
        d1 = int((ID%1000)/100)
        d2 = int((ID%100)/10)
        d3 = (ID%10)
        
        print("Solution for Q1:\n")
        A = set(generate_unique_list(d1, d2, d3, 1))
        B = set(generate_unique_list(d1, d2, d3, 2))
        X = A.union(B)
        print("A U B is: ", end='')
        print(X)
        Y = A.difference(B)
        print("A - B is: ", end ='')
        print(Y)
        ans = X.symmetric_difference(Y)
        print("\nSymmetric Difference is: ", end='')
        print(ans)
        
        
        print("\n\n\nSolution for Q2:\n")
        U = set([1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20])
        A = set(generate_unique_list(d1, d2, d3, 3))
        B = set(generate_unique_list(d1, d2, d3, 4))
        X = U.difference(A)
        print("A' is: ", end='')
        print(X)
        Y = U.difference(B)
        print("B' is: ", end ='')
        print(Y)
        ans = X.difference(Y)
        print("\nA'-B' is: ", end='')
        print(ans)
        
        
        




        # Write the updated content back to the file
        with open('file.txt', 'a') as file:
            file.write(file_content)

        print('\n\n\n')
    except FileNotFoundError:
        print(f'The file "{file_path}" was not found.')
    except Exception as e:
        print(f'An error occurred: {str(e)}')



print("\nEnter Student ID: ", end='')
ID = int(input())
Name = ""
for i in range(44):
    row_number = 3+i
    Id = int(sheet.cell_value(row_number, 2))
    if(Id == ID):
        Name = sheet.cell_value(row_number, 3)
        break
    
if(Name == ""):
    print("\nID not found.\n\n\n")
    exit()
    
print("\nStudent Name is: " + Name + "\n\n")
add_text_in_file(ID, Name)


