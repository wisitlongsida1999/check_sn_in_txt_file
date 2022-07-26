import os 
import xlwings as xw




files_ls = os.listdir('input')

wb = xw.Book('Check Known issue.xlsx') 
sheet = wb.sheets['Sheet1']

for file in files_ls:
    
    f = open(f'input\\{file}', "r")
    text = f.read()
    row = 2
    print('Search File : ',file)
    while True:
        sn = str(sheet.range(f'A{row}').value)
        file_name = str(sheet.range(f'B{row}').value)
        
        if sn in ['None','']:
            
            break
        
        if sn in text:
            
            if file_name in ['None','']:
            
                sheet.range(f'B{row}').value = file
                print('>>> Found : ',sn)
                
            elif file_name not in sheet.range(f'B{row}').value:

                sheet.range(f'B{row}').value = file_name+', '+file
                print('>>> Found : ',sn)
                
            else:
                
                print('>>> Already exists : ',sn)
                
        row+=1
    

    f.close()
    
wb.save()
