import pandas as pd
import os

def path_system():
    THIS_FOLDER = os.path.dirname(os.path.abspath(__file__)) + "/files"
    my_file = os.path.join(THIS_FOLDER, 'sat_final.xlsx')
    return my_file

def get_sheets():
    xls = pd.ExcelFile(path_system())
    sheets = xls.sheet_names
    get_dates_xls =  pd.read_excel(path_system(), sheet_name=sheets[0], dtype={'Domicilio Fiscal(C.P.):': str})

    #Imprimir las cabeceras
    #cabeceras = get_dates_xls.columns.tolist()

    #for i in cabeceras:
    #    get_dates_xls[i] = get_dates_xls[i].fillna(method='ffill')    

    get_dates_xls['ID'] = get_dates_xls['ID'].fillna(method='ffill')
    get_dates_xls['*Régimen Fiscal:'] = get_dates_xls['*Régimen Fiscal:'].fillna(method='ffill')
    get_dates_xls['*RFC:'] = get_dates_xls['*RFC:'].fillna(method='ffill')
    get_dates_xls['*Razón Social'] = get_dates_xls['*Razón Social'].fillna(method='ffill')
    get_dates_xls['Correo:'] = get_dates_xls['Correo:'].fillna(method='ffill')
    get_dates_xls['Domicilio Fiscal(C.P.):'] = get_dates_xls['Domicilio Fiscal(C.P.):'].fillna(method='ffill')
    get_dates_xls['Regimen_Fiscal'] = get_dates_xls['Regimen_Fiscal'].fillna(method='ffill')
    get_dates_xls['Customer Number'] = get_dates_xls['Customer Number'].fillna(method='ffill')
    get_dates_xls['Traslados Base IVA 16'] = get_dates_xls['Traslados Base IVA 16'].fillna(method='ffill')
    get_dates_xls['Income Amount SUM 1'] = get_dates_xls['Income Amount SUM 1'].fillna(method='ffill')
    get_dates_xls['Traslados Impuestos IVA 16:'] = get_dates_xls['Traslados Impuestos IVA 16:'].fillna(method='ffill')
    get_dates_xls['*Fecha Pago:'] = get_dates_xls['*Fecha Pago:'].fillna(method='ffill')
    get_dates_xls['Moneda P.'] = get_dates_xls['Moneda P.'].fillna(method='ffill')
    get_dates_xls['*Forma Pago P'] = get_dates_xls['*Forma Pago P'].fillna(method='ffill')
    get_dates_xls['Tipo Cambio P:'] = get_dates_xls['Tipo Cambio P:'].fillna(method='ffill')
    get_dates_xls['*Monto:'] = get_dates_xls['*Monto:'].fillna(method='ffill')


    grouped = get_dates_xls.groupby(['ID',
                                    '*Régimen Fiscal:',
                                    '*RFC:',
                                    '*Razón Social',
                                    'Correo:',
                                    'Domicilio Fiscal(C.P.):',
                                    'Regimen_Fiscal',
                                    'Customer Number',
                                    'Traslados Base IVA 16',
                                    'Income Amount SUM 1',
                                    'Traslados Impuestos IVA 16:',
                                    '*Fecha Pago:',
                                    'Moneda P.',
                                    '*Forma Pago P',
                                    'Tipo Cambio P:',
                                    '*Monto:'
                                    ]).agg({'SubT_Linea': list,
                                            '*Tipo Factor P:': list,
                                            'Iva_Linea': list,
                                            '*Impuesto P': list,
                                            'Tasa o Cuota P:': list,
                                            '*Id Doc.:': list,
                                            'serie_dr': list,
                                            '*Moneda DR:': list,
                                            '*Núm. Parcialidad': list,
                                            '*Saldo Anterior.:': list,
                                            '*Objeto Imp DR:': list,
                                            'Folio': list,
                                            'Equivalencia:': list,
                                            '*Imp. Pagado': list,
                                            '*Saldo Insoluto:': list}).reset_index()
    

    factura = []
    for _, row in grouped.iterrows():
        dic = {'ID':row['ID'],
                '*Régimen Fiscal:':row['*Régimen Fiscal:'],
                '*RFC:':row['*RFC:'],
                '*Razón Social':row['*Razón Social'],
                'Correo:':row['Correo:'],
                'Domicilio Fiscal(C.P.):':row['Domicilio Fiscal(C.P.):'],
                'Regimen_Fiscal':row['Regimen_Fiscal'],
                'Customer Number':row['Customer Number'],
                'Traslados Base IVA 16':row['Traslados Base IVA 16'],
                'Income Amount SUM 1':row['Income Amount SUM 1'],
                'Traslados Impuestos IVA 16:':row['Traslados Impuestos IVA 16:'],
                '*Fecha Pago:':row['*Fecha Pago:'],
                'Moneda P.':row['Moneda P.'],
                '*Forma Pago P':row['*Forma Pago P'],
                'Tipo Cambio P:':row['Tipo Cambio P:'],
                '*Monto:':row['*Monto:'],
                'SubT_Linea':row['SubT_Linea'],
                '*Tipo Factor P:':row['*Tipo Factor P:'],
                'Iva_Linea':row['Iva_Linea'],
                '*Impuesto P':row['*Impuesto P'],
                'Tasa o Cuota P:':row['Tasa o Cuota P:'],
                '*Id Doc.:':row['*Id Doc.:'],
                'serie_dr':row['serie_dr'],
                '*Moneda DR:':row['*Moneda DR:'],
                '*Núm. Parcialidad':row['*Núm. Parcialidad'],
                '*Saldo Anterior.:':row['*Saldo Anterior.:'],
                '*Objeto Imp DR:':row['*Objeto Imp DR:'],
                'Folio':row['Folio'],
                'Equivalencia:':row['Equivalencia:'],
                '*Imp. Pagado':row['*Imp. Pagado'],
                '*Saldo Insoluto:':row['*Saldo Insoluto:']}
        factura.append([{'ID':dic['ID']},
                        {'*Régimen Fiscal:':dic['*Régimen Fiscal:']},
                        {'*RFC:':dic['*RFC:']},
                        {'*Razón Social':dic['*Razón Social']},
                        {'Correo:':dic['Correo:']},
                        {'Domicilio Fiscal(C.P.):':dic['Domicilio Fiscal(C.P.):']},
                        {'Regimen_Fiscal':dic['Regimen_Fiscal']},
                        {'Customer Number':dic['Customer Number']},
                        {'Traslados Base IVA 16':dic['Traslados Base IVA 16']},
                        {'Income Amount SUM 1':dic['Income Amount SUM 1']},
                        {'Traslados Impuestos IVA 16:':dic['Traslados Impuestos IVA 16:']},
                        {'*Fecha Pago:':dic['*Fecha Pago:']},
                        {'Moneda P.':dic['Moneda P.']},
                        {'*Forma Pago P':dic['*Forma Pago P']},
                        {'Tipo Cambio P:':dic['Tipo Cambio P:']},
                        {'*Monto:':dic['*Monto:']},
                        {'SubT_Linea':dic['SubT_Linea']},
                        {'*Tipo Factor P:':dic['*Tipo Factor P:']},
                        {'Iva_Linea':dic['Iva_Linea']},
                        {'*Impuesto P':dic['*Impuesto P']},
                        {'Tasa o Cuota P:':dic['Tasa o Cuota P:']},
                        {'*Id Doc.:':dic['*Id Doc.:']},
                        {'serie_dr':dic['serie_dr']},
                        {'*Moneda DR:':dic['*Moneda DR:']},
                        {'*Núm. Parcialidad':dic['*Núm. Parcialidad']},
                        {'*Saldo Anterior.:':dic['*Saldo Anterior.:']},
                        {'*Objeto Imp DR:':dic['*Objeto Imp DR:']},
                        {'Folio':dic['Folio']},
                        {'Equivalencia:':dic['Equivalencia:']},
                        {'*Imp. Pagado':dic['*Imp. Pagado']},
                    {'*Saldo Insoluto:':dic['*Saldo Insoluto:']}])

    
    return factura





def delete_nan(list_factura):
    new_list = [item for item in list_factura if not(pd.isnull(item)) == True]
    return new_list



factura = 17
lista_factura = get_sheets()[factura - 1][0]['ID']

lista_impuestos_p_traslados = [delete_nan(get_sheets()[factura - 1][16]['SubT_Linea']),
                                delete_nan(get_sheets()[factura - 1][17]['*Tipo Factor P:']),
                                delete_nan(get_sheets()[factura - 1][18]['Iva_Linea']),
                                delete_nan(get_sheets()[factura - 1][19]['*Impuesto P']),
                                delete_nan(get_sheets()[factura - 1][20]['Tasa o Cuota P:']),
                                delete_nan(get_sheets()[factura - 1][21]['*Id Doc.:']),
                                delete_nan(get_sheets()[factura - 1][22]['serie_dr']),
                                delete_nan(get_sheets()[factura - 1][23]['*Moneda DR:']),
                                delete_nan(get_sheets()[factura - 1][24]['*Núm. Parcialidad']),
                                delete_nan(get_sheets()[factura - 1][25]['*Saldo Anterior.:']),
                                delete_nan(get_sheets()[factura - 1][26]['*Objeto Imp DR:']),
                                delete_nan(get_sheets()[factura - 1][27]['Folio']),
                                delete_nan(get_sheets()[factura - 1][28]['Equivalencia:']),
                                delete_nan(get_sheets()[factura - 1][29]['*Imp. Pagado']),
                                delete_nan(get_sheets()[factura - 1][30]['*Saldo Insoluto:']),]

                                     
                                   
#print(f"factura: {lista_factura}")
#print(get_sheets()[factura][5]['Domicilio Fiscal(C.P.):'])



if __name__ == '__main__':
    get_sheets()
    