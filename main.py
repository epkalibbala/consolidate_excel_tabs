import numpy as np
import pandas as pd
from datetime import datetime
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
        period = sheet_name.split('-')[-1]
        # sheet['PERIOD'] = sheet_name  # f'{period[-2:]}-{period[:-3].capitalize()}'
        if f'{period.strip()[:-3].capitalize()}' != 'Sept' and f'{period.strip()[:-2].capitalize()}' != 'July' and f'{period.strip()[:-3].capitalize()}' != 'Sep' and f'{period.strip()[:-2].capitalize()}' != 'Cot' and f'{period.strip()[:-2].capitalize()}' != 'June' and f'{period.strip()[:-2].capitalize()}' != 'Febe':
            datetime_object = datetime.strptime(f'{period.strip()[:-2].capitalize()} 20{period.strip()[-2:]}', '%b %Y')
        elif f'{period.strip()[:-3].capitalize()}' == 'July':
            datetime_object = datetime.strptime(f'Jul 20{period.strip()[-3:]}', '%b %Y')
        elif f'{period.strip()[:-2].capitalize()}' == 'Cot':
            datetime_object = datetime.strptime(f'Oct 20{period.strip()[-2:]}', '%b %Y')
        elif f'{period.strip()[:-2].capitalize()}' == 'Febe':
            datetime_object = datetime.strptime(f'Feb 20{period.strip()[-2:]}', '%b %Y')
        elif f'{period.strip()[:-2].capitalize()}' == 'June':
            datetime_object = datetime.strptime(f'Jun 20{period.strip()[-2:]}', '%b %Y')
        else:
            datetime_object = datetime.strptime(f'Sep 20{period.strip()[-2:]}', '%b %Y')

        date_time = pd.to_datetime(datetime_object, format='%m/%Y')

        edited_period = date_time.to_period('m')

        sheet['PERIOD'] = edited_period

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
# os.remove(f'C:/Users/emmanuelk/Desktop/excel/final/{file_name} consolidated.xlsx')  # Delete after use
os.rename(f'C:/Users/emmanuelk/Desktop/excel/final/{file_name}.xlsx',
          f'C:/Users/emmanuelk/Desktop/excel/final/{file_name} consolidated.xlsx')

# Reset your index or you'll have duplicates
df = df.reset_index(drop=True)

# df = \
#     pd.concat([df.assign(file=os.path.splitext(os.path.basename(f))[0], sheet=sheet) for f in glob(f_mask)
#                for sheet, df in pd.read_excel(f, sheet_name=None).items()],
#               ignore_index=True)
