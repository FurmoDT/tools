import pandas as pd
import os

filename = input('파일 이름(ex.. 면역, 위대한 여정 1부.xls): ')

# read
if filename.endswith('.xls'):
    df = pd.read_excel(filename, sheet_name=0, engine='xlrd')
elif filename.endswith('.xlsx'):
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
column = df.columns
# read

# split
value = df.values.tolist()
num = 1
TC_IN, TC_OUT, KOR, ENG = '', '', '', ''
new_list = list()
temp = [num, TC_IN, TC_OUT, KOR, ENG]
for i in value:
    if not temp[1]:
        temp[1] = i[1]
    temp[3] += ' ' + i[3].replace('|', ' ').rstrip()
    temp[4] += ' ' + i[4].replace('|', ' ').rstrip()
    if i[-1].endswith(('.', '?', '!')):
        temp[2] = i[2]
        temp[3] = temp[3].lstrip()
        temp[4] = temp[4].lstrip()
        new_list.append(temp)
        num += 1
        temp = [num, TC_IN, TC_OUT, KOR, ENG]
# split

# write
xlsx_dir = os.path.join(os.getcwd(), 'sentence.xlsx')

df = pd.DataFrame(new_list, columns=column)
df.to_excel(excel_writer=xlsx_dir,
            startrow=0,
            startcol=0,
            index=False)
# write
