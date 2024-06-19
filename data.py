import openpyxl as oxl

# Def. files
INPUT_file ='input.xlsx'
wb_input = oxl.load_workbook(INPUT_file)
OUTPUT_file = 'output.xlsx'
wb_output = oxl.load_workbook(OUTPUT_file)
# Def. Sheets
sheets = range(5)
sheets = [str(sheet) for sheet in sheets]
# Def. rows_I

# Def. rows_O
