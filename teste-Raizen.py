import pandas as pd

import numpy as np

import datetime


sheet_id = '1-DarzslxY6CndyoUr1y_z68YaX7Cvj-L'

# First, is important study the dataset and structure all information for Python. In this case, one can not
# extract information from a Excel Pivot Table. Therefore, it is necessary show details of the Pivot Table
# in Excel and organize all information in a dataset to prepare for Python

data = pd.read_excel (f'https://drive.google.com/file/d/{sheet_id}/export?format=xls',sheet_name='Sheet1') # Read the worksheet

dframe = pd.DataFrame(data); # Transform the worksheet in a Data Frame. It is important standardize data

dframe = dframe.drop('TOTAL', axis = 1); # Remove the last column (TOTAL) because the table format does not include total rows

dframe = dframe.rename(columns = {'Jan':'January', 'Fev':'February',
                                'Mar':'March', 'Abr':'April',
                                'Mai':'May', 'Jun':'June',
                                'Jul':'July', 'Ago':'August',
                                'Set':'September', 'Out':'October',
                                'Nov':'November', 'Dez':'December'}) # Months are set in portuguese, but Python datetime reads only date in english.


abb_states = {'ACRE': 'AC', 'ALAGOAS': 'AL','AMAPÁ': 'AP',
                              'AMAZONAS': 'AM', 'BAHIA': 'BA', 'CEARÁ': 'CE',
                              'DISTRITO FEDERAL': 'DF', 'ESPÍRITO SANTO': 'ES',
                              'GOIÁS': 'GO', 'MARANHÃO': 'MA', 'MATO GROSSO': 'MT',
                              'MATO GROSSO DO SUL': 'MS', 'MINAS GERAIS': 'MG',
                              'PARÁ': 'PA', 'PARAÍBA': 'PB', 'PARANÁ': 'PR',
                              'PERNAMBUCO': 'PE', 'PIAUÍ': 'PI', 'RIO DE JANEIRO': 'RJ',
                              'RIO GRANDE DO NORTE': 'RN', 'RIO GRANDE DO SUL': 'RS',
                              'RONDÔNIA': 'RO', 'RORAIMA': 'RR', 'SANTA CATARINA': 'SC',
                              'SÃO PAULO': 'SP', 'SERGIPE': 'SE', 'TOCANTINS': 'TO'} # Setting the pairs of names and abbreviation of each state

dframe['uf'] = dframe['ESTADO'].map(abb_states) # Creating the column 'uf' according to the state name


dframe = dframe.melt(id_vars=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE','uf'],
                     var_name = 'MONTH', value_name = 'VOLUME'); # Transposing month columns from variables to observables

dframe['created_at'] = pd.to_datetime(dframe['ANO'].astype(str)  + dframe['MONTH'].astype(str), format='%Y%B') # Adding column 'created_at' based datetime

dframe['year_month'] = dframe['created_at'].dt.strftime('%Y-%m') # Creating 'year_month' converting datetime from 'created_at'

dframe.VOLUME = dframe.VOLUME.round(2); # Standardising the values for double precision


oil_dev = dframe[(dframe.COMBUSTÍVEL == 'ÓLEO COMBUSTÍVEL (m3)') | (dframe.COMBUSTÍVEL == 'GASOLINA DE AVIAÇÃO (m3)') |
                  (dframe.COMBUSTÍVEL == 'QUEROSENE ILUMINANTE (m3)') | (dframe.COMBUSTÍVEL == 'QUEROSENE DE AVIAÇÃO (m3)') |
                  (dframe.COMBUSTÍVEL == 'ÓLEO DIESEL (m3)') | (dframe.COMBUSTÍVEL == 'GLP (m3)') |
                  (dframe.COMBUSTÍVEL == 'GASOLINA C (m3)')]; # Create first table filtering wished product

diesel = dframe[(dframe.COMBUSTÍVEL == 'ÓLEO DIESEL (m3)')]; # Create second table filtering wished product

oil_dev = oil_dev.rename(columns = {'COMBUSTÍVEL':'product','UNIDADE':'unit','VOLUME':'volume'}); # Renaming the columns

oil_dev = oil_dev[['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']]; # Putting columns in desired order

diesel = diesel.rename(columns = {'COMBUSTÍVEL':'product','UNIDADE':'unit','VOLUME':'volume'}); # Renaming the columns

diesel = diesel[['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']]; # Putting columns in desired order

print(oil_dev)

print(diesel)

