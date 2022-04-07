import numpy as np
import pandas as pd
import os
from glob import glob

file_name = 'BAYPORT BATCH 0158 (002)'
path = fr'I:/CREDIT/CREDIT ADMIN/PCA Batch Schedules/{file_name}.xlsx'
workbook = pd.read_excel(path, sheet_name=None)
# workbook.popitem()
df = pd.DataFrame()
work_dict = workbook.items()
for sheet_name, sheet in work_dict:
    if sheet_name != "SUMMARY":
        period = sheet_name.split('-')[1]
        sheet['PERIOD'] = period
        print(sheet_name.split('-')[1])
        workbook[sheet_name] = sheet
        print(workbook[sheet_name].head())
        df = df.append(sheet)

# conso = pd.concat(workbook, ignore_index=True)
# print(conso)
# conso.to_csv(r'C:/Users/emmanuelk/PycharmProjects/consolidate/export_dataframe.csv', index=False)
df['EMPLOYEE'].replace('', np.nan, inplace=True)
df.dropna(subset=['EMPLOYEE'], inplace=True)
with pd.ExcelWriter(f'C:/Users/emmanuelk/Desktop/excel/final/{file_name}.xlsx',
                    mode='a') as writer:
    df.to_excel(writer, sheet_name='consolidation', index=False)
    # pd.read_csv(r'C:/Users/emmanuelk/Desktop/excel/export_dataframe.csv').to_excel(writer, sheet_name='consolidation')
df.to_csv(fr'C:/Users/emmanuelk/Desktop/excel/final/csv/{file_name} consolidated.csv', index=False)
os.rename(f'C:/Users/emmanuelk/Desktop/excel/final/{file_name}.xlsx', f'C:/Users/emmanuelk/Desktop/excel/final/{file_name} consolidated.xlsx')

# Reset your index or you'll have duplicates
df = df.reset_index(drop=True)

# df = \
#     pd.concat([df.assign(file=os.path.splitext(os.path.basename(f))[0], sheet=sheet) for f in glob(f_mask)
#                for sheet, df in pd.read_excel(f, sheet_name=None).items()],
#               ignore_index=True)
