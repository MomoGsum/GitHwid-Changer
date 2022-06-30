# Importing prerequisites

from tkinter.tix import InputOnly
from isort import file
import pandas as pd
from sympy import inverse_laplace_transform
from pathlib import Path

# Setting up file directory in which the dataset contains

folder = Path('/Users/rian/Documents/Codes/Latihan/DataAnalytic/Tietex/JOIN')

#Preparing list for dataset to merge
list = []

# creating a function to clean the dataset
def clean(input_file_directory):
    name = Path(input_file_directory).stem
    df = pd.read_excel(input_file_directory)

    # Renaming column and column name typos
    df.columns = [col.upper() for col in df.columns]
    df.columns = [col.strip() for col in df.columns]
    df.columns = [col.replace('"', '') for col in df.columns]
    for i in df.columns:
        if 'ASED' in i:
            df.rename(columns={i: 'DATE'}, inplace=True)
    for i in df.columns:
        if 'RELEA' in i:
            df.rename(columns={i: 'DATE'}, inplace=True)
    for i in df.columns:
        if 'CUST' in i:
            df.rename(columns={i: 'CUSTOMER'}, inplace=True)

    # Delete 'No' column because we won't need it
    if 'NO' in df.columns:
        df = df.drop('NO', axis=1)
    
    # Uppercasing all customer name and filling NA with name of value above
    df['CUSTOMER'] = df['CUSTOMER'].str.upper()
    df['CUSTOMER'].fillna(method='ffill', inplace=True)

    # Filling na of 'DATE' column with the value of 'DATE' column above and changing type to 'str'
    df['DATE'].fillna(method='ffill', inplace=True)
    df['DATE'].astype('str')

    # Filling NA in item quantity with 0
    df.fillna(0, inplace=True)

    # Melting dataset from wide to long
    df_fixed = df.melt(id_vars=['CUSTOMER','DATE'], var_name='ITEM', value_name='QUANTITY')
    
    # Uppercasing every item name and dropping duplicates
    df_fixed['ITEM'] = df_fixed['ITEM'].str.upper()
    df_fixed.drop_duplicates(inplace=True)

    # Renaming and grouping item into categories
    for a in df_fixed['ITEM']:
        if '45' in a:
            df_fixed.loc[df_fixed['ITEM'] == a, 'ITEM'] = 'COLOR'
    for b in df_fixed['ITEM']:
        if 'RGS' in b:
            df_fixed.loc[df_fixed['ITEM'] == b, 'ITEM'] = 'RGS'
    for c in df_fixed['ITEM']:
        if 'COLOR' in c:
            df_fixed.loc[df_fixed['ITEM'] == c, 'ITEM'] = 'COLOR'
    for d in df_fixed['ITEM']:
        if 'BLACK' in d:
            df_fixed.loc[df_fixed['ITEM'] == d, 'ITEM'] = 'PLM300(BLACK)'
    for e in df_fixed['ITEM']:
        if 'LINING MATERIAL' in e:
            df_fixed.loc[df_fixed['ITEM'] == e, 'ITEM'] = 'PLM300(WHITE)'
    for f in df_fixed['ITEM']:
        if f not in ['COLOR','RGS','PLM300(BLACK)','PLM300(WHITE)']:
            df_fixed.loc[df_fixed['ITEM'] == f, 'ITEM'] = 'COLOR'

    # Stripping and uppercasing customer name
    df['CUSTOMER'] = df['CUSTOMER'].str.upper()
    df['CUSTOMER'] = df['CUSTOMER'].str.strip()

    # Renaming customer name so that it has consistent name
    for customer1 in df_fixed['CUSTOMER']:
        if 'CHANG' in customer1:
            df_fixed.loc[df_fixed['CUSTOMER'] == customer1, 'CUSTOMER'] = 'CHANGSHIN'
    for customer2 in df_fixed['CUSTOMER']:
        if 'IKOMAS' in customer2:
            df_fixed.loc[df_fixed['CUSTOMER'] == customer2, 'CUSTOMER'] = 'NIKOMAS'
    for customer3 in df_fixed['CUSTOMER']:
        if 'POU' in customer3:
            df_fixed.loc[df_fixed['CUSTOMER'] == customer3, 'CUSTOMER'] = 'POUYUEN'
    for customer4 in df_fixed['CUSTOMER']:
        if 'TK.' in customer4:
            df_fixed.loc[df_fixed['CUSTOMER'] == customer4, 'CUSTOMER'] = 'TK'
    for customer5 in df_fixed['CUSTOMER']:
        if 'SELALU' in customer5:
            df_fixed.loc[df_fixed['CUSTOMER'] == customer5, 'CUSTOMER'] = 'SCI'

    # Making sure there is no NA in 'CUSTOMER' and 'ITEM' column
    df_fixed = df_fixed[df_fixed['CUSTOMER'].notna()]
    df_fixed = df_fixed[df_fixed['ITEM'] != '']

    # appending cleaned data set into list and merging all data set into one
    list.append(df_fixed)

    datareal = pd.concat(list)
    datareal.drop_duplicates(inplace=True)

    # Excluding total order released from dataset
    datareal = datareal[datareal['CUSTOMER'] != 'TOTAL ORDER RELEASED (OIA)']

    # Exporting cleaned dataset into csv on specific path
    datareal.to_csv(f'/Users/rian/Documents/Codes/Latihan/DataAnalytic/Tietex/SCRIPT/output/cleaned.csv', index=False)

# Calling the function to clean the dataset
for file in folder.iterdir():
    clean(file)