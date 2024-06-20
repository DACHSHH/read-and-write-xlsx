import openpyxl as oxl
def get_number_of_rows(file, sheet):
    # Load the workbook and select the sheet
    
    
    # Initialize the count
    count = 0
    # Iterate over the rows
    for row in file[sheet].iter_rows(min_col=3,max_col=3):
        # If the first cell in the row is not empty, increment the count
        if row[0].value is not None:
            count += 1
            # print(str(sheet) +','+ str(count))
    # Return the count
    return count

def is_EPIC(string):
    # Zähle die Anzahl der Punkte im String
    dot_count = string.count('.')
    # Wenn der String genau einen oder zwei Punkte enthält, gib True zurück. Ansonsten gib False zurück.
    return dot_count == 1