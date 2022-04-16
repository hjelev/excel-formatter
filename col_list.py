import string

def column_letters():
    col_range = list(string.ascii_lowercase)
    new_list = list(string.ascii_lowercase)

    for c in col_range:
        for letter in col_range:
            new_list.append(c + letter)
    return new_list
    
print(column_letters())