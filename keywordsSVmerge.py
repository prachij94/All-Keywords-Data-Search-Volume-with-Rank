# -*- coding: utf-8 -*-
"""
Created on Wed Sep 19 11:51:11 2018

@author: Prachi Jain
"""

import pandas as pd


#Reading the input data
df1 = pd.read_excel("C:/Users/IMART/Downloads/GA Mob (1).xlsx",sheet_name = "All GA Keywords",encoding = "ISO-8859-1")

df1 = df1.drop('Attribute',axis=1)

df1.rename(columns={'Keyword': 'Keywords'}, inplace=True)

df2 = pd.read_excel("C:/Users/IMART/Downloads/webmaster (1).xlsx",sheet_name = "Export Worksheet",encoding = "ISO-8859-1")

df2 = df2.drop('Attribute',axis=1)

df2.rename(columns={'FK_GLCAT_MCAT_ID': 'MCAT ID','GLCAT_MCAT_NAME':'MCAT Name','TOT_KEYWORD_WEBMASTER':'Keywords'}, inplace=True)

df3 = pd.read_excel("C:/Users/IMART/Downloads/mob_final_data (1).xlsx",sheet_name = "mob_final_data-Tableau",encoding = "ISO-8859-1")

df3 = df3.drop('Attribute',axis=1)

df3.rename(columns={'MCAT Id': 'MCAT ID','KW':'Keywords'}, inplace=True)

df4 = pd.read_excel("C:/Users/IMART/Downloads/Product Groups (1).xlsx",sheet_name = "Sheet1",encoding = "ISO-8859-1")

df4 = df4.drop('Attribute',axis=1)

df4.rename(columns={'MCAT id': 'MCAT ID','Keyword': 'Keywords'}, inplace=True)

df5 = pd.read_excel("C:/Users/IMART/Downloads/Product Groups (1).xlsx",sheet_name = "Sheet2",encoding = "ISO-8859-1")

df5 = df5.drop('Attribute',axis=1)

df5.rename(columns={'MCAT id': 'MCAT ID','Keyword': 'Keywords'}, inplace=True)

df6 = pd.read_excel("C:/Users/IMART/Downloads/Product Groups (1).xlsx",sheet_name = "Sheet3",encoding = "ISO-8859-1")

df6 = df6.drop('Attribute',axis=1)

df6.rename(columns={'MCAT id': 'MCAT ID','Keyword': 'Keywords'}, inplace=True)

df7 = pd.read_excel("C:/Users/IMART/Downloads/ExtraKeywords (1).xlsx",sheet_name = "ExtraKeywords",encoding = "ISO-8859-1")

df7 = df7.drop('source',axis=1)

df7.rename(columns={'MCAT id': 'MCAT ID','Keyword': 'Keywords'}, inplace=True)

df7 = df7[['MCAT ID','MCAT Name','Keywords']]

inputdataframe = pd.concat(f for f in [df1,df2,df3,df4,df5,df6,df7])

inputdataframe.drop_duplicates(subset=['MCAT ID', 'Keywords'], keep=False,inplace=True)

#Search volume data

df8 = pd.read_csv("C:/Users/IMART/Downloads/searchvolume.csv")

df8 = df8[['Related Keywords','Search Volume']]

df8.rename(columns={'Related Keywords': 'Keywords'}, inplace=True)

df9 = pd.read_csv("C:/Users/IMART/Downloads/SV-Output1.csv")

df9 = df9[['keyword','searchvolume']]

df9.rename(columns={'keyword': 'Keywords','searchvolume':'Search Volume'}, inplace=True)

df10 = pd.read_csv("C:/Users/IMART/Downloads/SV-Output2.csv")

df10 = df10[['keyword','searchvolume']]

df10.rename(columns={'keyword': 'Keywords','searchvolume':'Search Volume'}, inplace=True)

df11 = pd.read_csv("C:/Users/IMART/Downloads/SV-Output3.csv")

df11 = df11[['keyword','searchvolume']]

df11.rename(columns={'keyword': 'Keywords','searchvolume':'Search Volume'}, inplace=True)

df12 =  pd.read_excel("C:/Users/IMART/Downloads/originalKeywordonlyNA.xlsx",sheet_name="originalKeywordonlyNA")

df12 = df12[['keyword','searchvolume']]

df12.rename(columns={'keyword': 'Keywords','searchvolume':'Search Volume'}, inplace=True)

df13 = pd.read_excel("C:/Users/IMART/Downloads/originalKeywordonlyNA.xlsx",sheet_name="Sheet1")

df13.rename(columns={'Keyword': 'Keywords','Search Volume (IN)':'Search Volume'}, inplace=True)

searchvolumedataframe = pd.concat(f for f in [df8,df9,df10,df11,df12,df13])

searchvolumedataframe.drop_duplicates(subset=['Keywords','Search Volume'], keep=False,inplace=True)


result = inputdataframe.merge(searchvolumedataframe,how='left',on='Keywords')

result.to_csv("FinalSearchVolumewithMCAT.csv",index=False)

df14 = pd.read_csv("C:/Users/IMART/Downloads/Subcat_mcat.csv",encoding = "ISO-8859-1")

df14.rename(columns={'GLCAT_MCAT_ID': 'MCAT ID','FK_GLCAT_CAT_ID':'SUBCAT ID','GLCAT_CAT_NAME':'SUBCAT NAME'}, inplace=True)

final = result.merge(df14,how='left',on='MCAT ID')

final = final.drop('GLCAT_MCAT_NAME',axis=1)

final.to_csv("FinalSearchVolumewithMCATandSUBCATs.csv",index=False)