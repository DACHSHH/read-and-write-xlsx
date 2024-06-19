from functions import *
from data import *
# for sheet in sheets in INPUT file
for sheet in sheets:
     # for row_I in rows_I
    for row in range(get_number_of_rows(wb_input, sheet)):
        
        # if Epic <=> if string in C has one "."
            # 	 "Epic" -> B
            # Öffnen Sie die Eingabedatei und lesen Sie den Wert
            sheet_input = wb_input[sheet]
            input_value = sheet_input['C' + str(row + 1)].value

            # Öffnen Sie die Ausgabedatei und schreiben Sie den Wert
            
            sheet_output = wb_output['Upload_Template']
            sheet_output['C' + str(row + 1)] = input_value
            # 	 F + " - M" + sheet] -> C
            # 	  "PIP Milestone" + sheet] -> I, K, L
        #  if Deliverable <=> if string in C has one "."
            # 	 "Deliverable" -> B
            # 	 C + " - " + F - > C
            # 	 G -> E # Einzeilig
            # 	 P -> H # Ausgeschrieben
            # 	 R -> G # Ausgeschrieben
            # 	 "PIP Milestone" + sheet(Epic) -> L


# Speichern Sie die Änderungen in der Ausgabedatei
wb_output.save(OUTPUT_file)


