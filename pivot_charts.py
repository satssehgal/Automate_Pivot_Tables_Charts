import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import os

file=input('Filepath: ')
piv_index=input('Write Your Index: ').split()
piv_data=input('Columns to Graph: ').split()
print(piv_data[-1])


extension = os.path.splitext(file)[1]
filename = os.path.splitext(file)[0]
pth=os.path.dirname(file)
newfile=os.path.join(pth,filename+'_2'+extension)

xl = pd.ExcelFile(file)
sheetnames=xl.sheet_names
file=[pd.read_excel(file, sheet_name=s) for s in sheetnames]


for k in range(len(sheetnames)):

    df1=pd.DataFrame(file[k])
    df_pt=pd.pivot_table(df1, index = piv_index,values= piv_data)
    img=df_pt.plot(kind='bar',secondary_y=piv_data[-1], width=0.8, figsize=(12,10), title=sheetnames[k])
    imgdata = BytesIO()
    fig = plt.figure()
    img.figure.savefig(imgdata)


    writer = pd.ExcelWriter('{}/{}.xlsx'.format(pth,sheetnames[k]), engine='xlsxwriter')
    df_pt.to_excel(writer, sheet_name=sheetnames[k])
    worksheet = writer.sheets['{}'.format(sheetnames[k])]
    worksheet.insert_image('F1','',{'image_data': imgdata})
    writer.save()