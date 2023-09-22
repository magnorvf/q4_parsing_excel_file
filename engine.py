import pandas as pd
import os

doc = pd.ExcelFile('Inconsistency Report1.xlsx')

sheet_names = doc.sheet_names
dt_output = []

for sheet in sheet_names: #ACESSANDO AS PLANILHAS

    if "Sheet1" in sheet: 
        dt_frame = doc.parse(sheet)

        #SLICING DATAFRAME

        dt_frame.columns = dt_frame.iloc[5]
        dt_frame = dt_frame.iloc[6:, :9]

        rows = dt_frame.__len__()

        dt_frame = dt_frame.fillna(0)

        if dt_output.__len__() == 0:
            head = list(dt_frame.columns)
            dt_output.append([head[0]]+head[5:])

        stor_dw = ''

        for row in range(0, rows):
            interv = dt_frame.iloc[row, 5:]
            dw_no = dt_frame.iloc[row, 0]
            
            if dw_no != 0:
                stor_dw = str(dw_no)

            dt_output.append([stor_dw] + list(interv))


dt_frame_out = pd.DataFrame(dt_output)
fileoutput = 'saida.xlsx'
index = 0
while os.path.isfile(fileoutput):
    fileoutput = fileoutput.replace('.xlsx', f'_{index}.xlsx')

dt_frame_out.to_excel(fileoutput, index=False, header=False)


                

            

            

