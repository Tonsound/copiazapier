import pandas as pd
from pandas import ExcelWriter



def checkeador_diferencias(df1, df2):
    if len(df2)>len(df1):
        diff = len(df2)-len(df1)
        df3 = df2[len(df2)-diff:len(df2)].reset_index(drop=True)
    else:
        df3 = []
    return df3

def generar_DF_Excel(df, nombre_archivo):
    nombre_final = nombre_archivo + '.xls' 
    writer = ExcelWriter(nombre_final)
    df.to_excel(writer,'Quiebres_de_stock')
    writer.save()
    print('Ok')

xls1 = pd.ExcelFile('Historicas.xlsx')
df1 = pd.read_excel(xls1, 'Hoja1')

xls2 = pd.ExcelFile('Etiquetas Notorious.xlsx')
df2 = pd.read_excel(xls2, 'Hoja1')

df3 = checkeador_diferencias(df1, df2)
# generar_DF_Excel(df2, 'Historicas')

csv1 = pd.read_csv('QUIEBREShistoricos.csv', sep=',')
csv2 = pd.read_csv('QUIEBRES DE STOCK PRESTASHOP-Grid view.csv', sep=',')
df4 = checkeador_diferencias(csv1, csv2)
# csv2.to_csv('QUIEBREShistoricos.csv', sep=',', index = False)

def acople_de_data(df1, df2):
    new_df2 = df2
    for i in range(len(new_df2)):
        df3 = df1.loc[df1['CODIGO DE BODEGA'] == str(new_df2.loc[i, 'Código'])].reset_index()
        if len(df3) > 0:
            df4 = new_df2.loc[new_df2['Código'] == str(df3.loc[0, 'CODIGO DE BODEGA'])]
            print(df4.index[0])
            new_df2.loc[df4.index[0],'Nombre' ] = df3.loc[0, 'DESTINATARIO']
            new_df2.loc[df4.index[0],'Correo' ] = df3.loc[0, 'EMAIL']
            new_df2.loc[df4.index[0],'Numero' ] = '56' + str(int(df3.loc[0, 'TELEFONO']))
            if str(new_df2.loc[df4.index[0], 'Sugerencia']) == 'nan':
                new_df2.loc[df4.index[0],'WhatsApp' ]  = str(new_df2.loc[df4.index[0],'Whatsapp sin cambio' ]).replace('phone=&','phone=' + str(new_df2.loc[df4.index[0],'Numero' ]) + '&').replace('Hola%20%20como', 'Hola%20'  + str(new_df2.loc[df4.index[0],'Nombre'])+ '%20como' )
            else:
                new_df2.loc[df4.index[0],'WhatsApp' ]  = str(new_df2.loc[df4.index[0],'WhatsApp con cambio']).replace('phone=&','phone=' + str(new_df2.loc[df4.index[0],'Numero' ]) + '&').replace('Hola%20%20como', 'Hola%20'  + str(new_df2.loc[df4.index[0],'Nombre'])+ '%20como' )  
    new_df2.drop(columns=['Whatsapp sin cambio', 'WhatsApp con cambio'],inplace=True)
    return new_df2




df5 = acople_de_data(df3, df4)

generar_DF_Excel(df5, 'Contactables')
