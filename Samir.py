import streamlit as st
import pandas as pd
import numpy as np
import xlsxwriter
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from datetime import datetime, timedelta
from pandas.tseries.offsets import *
from openpyxl import load_workbook
import matplotlib.pyplot as plt




st.write("""
# Windsor Locks Daily Team Load Report Builder
This app produces daily report for C&S Windsor Locks facility.
""")

st.sidebar.header('User Input Features')

st.sidebar.markdown("""
[Example Shipping File](https://github.com/swade-55/Robi/blob/main/Est.%20Shpg%20Volume%20by%20Order.csv?raw=true)
""")

# Collects user input features into dataframe
triceps_file = st.sidebar.file_uploader("Upload your input Shipping file", type=["csv"])



if triceps_file is not None:
    df = pd.read_csv(triceps_file)
    df = df[(df['INBOUND'].str.len() == 5)]
    df.columns = df.iloc[0]
    df = df[df.Order != 'Order']
    df1 = df.copy()
    df1.loc[df.Customer.str.len().eq(3),'Customer']=np.nan

    #df1 = df.loc[df['Customer'] > 1, 'Customer'] = 1
    df1[df1['Customer'].isnull()]=df1[df1['Customer'].isnull()].shift(-1,axis=1)
    df1 = df1.drop(columns=['Cube'])
    df1 = df1.drop(columns=['Weight'])
    df1 = df1.rename(columns={'Cases': 'Total Cube', 'Lines': 'Total Weight', 'Route': 'Total Cases', '#': 'Total Lines', 'Invoice': 'Stops', 'SCTN': 'Total Routes'})
    df1['Total Cases'] = df1['Total Cases'].astype(float)
    df1['Total Lines'] = df1['Total Lines'].astype(float)
    numbers = df1.groupby(['Total Routes'], as_index=False).sum()
    #numbers = numbers.drop(columns=['Order','Customer','Weight','Cube'])
    #numbers = numbers.sort_values(by=['Cases','Lines'], ascending=False)
    numbers.reset_index(drop=True, inplace=True)
    stops = df1[['Total Routes','Stops']]
    #stops = stops.groupby(['Total Routes'], as_index=False).count().unique()
    stops = stops.groupby('Total Routes')['Stops'].nunique()
    pivot = numbers.merge(stops, how='inner', left_on=['Total Routes'], right_on=['Total Routes'])
    #pivot = pivot.sort_values(by=['Total Routes'], ascending=True)
    #df2 = df2.rename(columns = {'Total Routes':'Routes'})
    #df2 = df2[df2['Routes'].notna()]

    # Function to save all dataframes to one single excel
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=True, sheet_name='Most Lines Per Route')
        workbook = writer.book
        #worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        #worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    df_xlsx = to_excel(pivot)
    st.download_button(label='📥 Download Current Result', data=df_xlsx ,file_name= 'Most_Lines_Per_Route.xlsx')

        

    st.subheader('Routes Ranked by Case Volume')
    st.write(pivot)