from functions import *
from data import *
# Öffnen Sie die Ausgabedatei und schreiben Sie den Wert
sheet_output = wb_output['Upload_Template']
# for sheet in sheets in INPUT file

for sheet in sheets:
    epic = 0
     # for row_I in rows_I
    for row in range(get_number_of_rows(wb_input, sheet)):
         # Öffnen Sie die Eingabedatei und lesen Sie den Wert
        sheet_input = wb_input[sheet]
        C_str = sheet_input['C' + str(row + 1)].value
        F_str = sheet_input['F' + str(row + 1)].value
        # if Epic <=> if string in C has one "."
        if is_EPIC(C_str):
            # "Epic" -> B
            sheet_output['B' + str(row + 2)] = 'Epic'
            # 	 F + " - M" + sheet -> C, K
            sheet_output['C' + str(row + 2)] = F_str + ' - M' + str(sheet)
            sheet_output['K' + str(row + 2)] = F_str + ' - M' + str(sheet)
            # "PIP Milestone" + sheet -> I
            sheet_output['I' + str(row + 2)] = 'PIP Milestone ' + str(sheet)
            epic += 1
        # if Deliverable <=> if string in C has one "."
        else:
            # "Deliverable" -> B
            sheet_output['B' + str(row + 2)] = 'Deliverable'
            # C + " - " + F - > C
            sheet_output['C' + str(row + 2)] = C_str + ' - ' + C_str
            # G -> E # Einzeilig
            G_str = sheet_input['G' + str(row + 1)].value
            sheet_output['E' + str(row + 2)] = G_str
            # P -> H # Ausgeschrieben
            P_str = sheet_input['P' + str(row + 1)].value
            sheet_output['H' + str(row + 2)] = P_str
            # R -> G # Ausgeschrieben
            R_str = sheet_input['R' + str(row + 1)].value
            sheet_output['G' + str(row + 2)] = R_str
            # "PIP Milestone" + sheet -> I
            sheet_output['I' + str(row + 2)] = 'PIP Milestone ' + str(sheet)
            # Epic Name -> L <=> F(epic) + " - M" + sheet -> L
            F_str_epic = sheet_input['F' + str(epic)].value
            sheet_output['L' + str(row + 2)] = F_str_epic + ' - M' + str(sheet)



# Speichern Sie die Änderungen in der Ausgabedatei
wb_output.save(OUTPUT_file)


