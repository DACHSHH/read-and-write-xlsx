import openpyxl as oxl
def get_number_of_rows(filename, sheetname):
    # Load the workbook and select the sheet
    wb = oxl.load_workbook(filename)
    sheet = wb[sheetname]
    # Initialize the count
    count = 0
    # Iterate over the rows
    for row in sheet.iter_rows():
        # If the first cell in the row is not empty, increment the count
        if row[0].value is not None:
            count += 1
    # Return the count
    return count

def is_EPIC(string):
    # Zähle die Anzahl der Punkte im String
    dot_count = string.count('.')
    
    # Wenn der String genau einen oder zwei Punkte enthält, gib True zurück. Ansonsten gib False zurück.
    return dot_count == 1 or dot_count == 2