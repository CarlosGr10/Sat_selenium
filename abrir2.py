import re
import pandas as pd
import pathlib
import os
import math
import numpy as np
from numpy import nan

def path_system():
    THIS_FOLDER = os.path.dirname(os.path.abspath(__file__)) + "/files"
    my_file = os.path.join(THIS_FOLDER, 'pais.xlsx')
    return my_file

def get_sheets():
    xls = pd.ExcelFile(path_system())
    sheets = xls.sheet_names
    get_dates_xls =  pd.read_excel(path_system(), sheet_name=sheets[0])
    print(get_dates_xls)

    cabeceras = get_dates_xls.columns.tolist()

    #Imprimir las cabeceras
    print(cabeceras)
    

    
    get_dates_xls['id'] = get_dates_xls['id'].fillna(method='ffill')
    get_dates_xls['Pais'] = get_dates_xls['Pais'].fillna(method='ffill')
    get_dates_xls['code_pais'] = get_dates_xls['code_pais'].fillna(method='ffill')

    #get_dates_xls['estados'] = get_dates_xls['estados'].fillna(method='ffill')


    # Agrupar por id y Pais, y aplicar la función de agregación a estados
    grouped = get_dates_xls.groupby(['id', 'Pais', 'code_pais']).agg({'estados': list, 'code_estado': list}).reset_index()
    print(grouped)
    # Construir la lista de diccionarios
    lista = []
    for _, row in grouped.iterrows():
        dic = {'id': int(row['id']), 'Pais': row['Pais'],'code_pais':row['code_pais'], 'estados': row['estados'], 'code_estado': row['code_estado']}
        lista.append([{'id': dic['id']}, {'Pais': dic['Pais']},{'code_pais': dic['code_pais']}, {'estados': dic['estados']}, {'code_estado': dic['code_estado']}])


    print(lista[5][4]['code_estado'])

    new_list = []
    for item in lista[5][4]['code_estado']:
        if str(item) != 'nan':
            new_list.append(item)
 
    print(new_list)

    new_list = [item for item in lista[0][4]['code_estado'] if not(pd.isnull(item)) == True]
    print(new_list)
    
if __name__ == '__main__':
    get_sheets()