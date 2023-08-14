from excel_methods import *



wb = load_workbook("uszips.xlsx")
ws = wb.active
states = []

# if there is no states folder, create one 

for i, row in enumerate(ws):
    
    dmm = []
    
    if i > 0:
        
    
        if i == 1:
            
            cs = ws[f"F{i+1}"].value
            states.append(cs)
            
            wbs = Workbook()
            wss = wbs.active
        
        
        
        if ws[f"F{i+1}"].value == cs:
            
            for cell in row:
                dmm.append(cell.value)
                
            wss.append(dmm)
            
            
        else:
            
            wbs.save(f'.\\States\\{cs}.xlsx')
            cs = ws[f"F{i+1}"].value
            
            
            # Check if a state file already exists
            if cs in states:
                
                wbs = load_workbook(f'.\\States\\{cs}.xlsx')
                wss = wbs.active
                
            else:
                
                
                wbs = Workbook()
                wss = wbs.active
            
            
            for cell in row:
                dmm.append(cell.value)
                
            wss.append(dmm)
            states.append(cs)
            
            
wbs.save(f'.\\States\\{cs}.xlsx')
        
        