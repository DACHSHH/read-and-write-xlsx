from functions import *
from data import *
# Öffnen Sie die Ausgabedatei und schreiben Sie den Wert
sheet_output = wb_output['Upload_Template']
# for sheet in sheets in INPUT file
row_output = 0
for sheet in sheets:
    epic = 0
     # for row_I in rows_I
    for row_input in range(get_number_of_rows(wb_input, sheet)):
         # Öffnen Sie die Eingabedatei und lesen Sie den Wert
        sheet_input = wb_input[sheet]
        C_str = sheet_input['C' + str(row_input + 1)].value
        F_str = sheet_input['F' + str(row_input + 1)].value
        row_output += 1
        # if Epic <=> if string in C has one "."
        if is_EPIC(C_str):
            # "Epic" -> B
            sheet_output['B' + str(row_output + 1)] = 'Epic'
            # 	 F + " - M" + sheet -> C, K
            sheet_output['C' + str(row_output + 1)] = F_str + ' - M' + str(sheet)
            sheet_output['K' + str(row_output + 1)] = F_str + ' - M' + str(sheet)
            # "TEP Milestone" + sheet -> I
            sheet_output['I' + str(row_output + 1)] = 'TEP Milestone ' + str(sheet)
            epic = row_input 
        # if Deliverable <=> if string in C has one "."
        else:
            # "Deliverable" -> B
            sheet_output['B' + str(row_output + 1)] = 'Deliverable'
            # C + " - " + F -> C
            sheet_output['C' + str(row_output + 1)] = C_str + ' - ' + F_str
            # G -> E # Einzeilig
            G_str = sheet_input['G' + str(row_input + 1)].value.replace('\n', '')
            sheet_output['E' + str(row_output + 1)] = G_str
            # P -> H # Ausgeschrieben
            P_str = sheet_input['P' + str(row_input + 1)].value
            sheet_output['H' + str(row_output + 1)] = P_str
            # R -> G # Ausgeschrieben
            R_str = sheet_input['R' + str(row_input + 1)].value
            sheet_output['G' + str(row_output + 1)] = R_str
            # "TEP Milestone" + sheet -> I
            sheet_output['I' + str(row_output + 1)] = 'TEP Milestone ' + str(sheet)
            # Epic Name -> L <=> F(epic) + " - M" + sheet -> L
            F_str_epic = sheet_input['F' + str(epic + 1)].value
            sheet_output['L' + str(row_output + 1)] = F_str_epic + ' - M' + str(sheet)



# Speichern Sie die Änderungen in der Ausgabedatei
wb_output.save(OUTPUT_file)


