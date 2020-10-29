#!/usr/bin/env python
# coding: utf-8


#Project name: Excel --> netCDF
#Description: It receives an excel file with a predefined format and returns the equivalent file in netCDF format. 
#Programmers: Amir \& Ashley
#Date: 07-09-2020



import pandas as pd
import xarray


print("Enter the exact name of the workbook. Please do include the file extension (e.g, .xlsx, .xls, etc):")
workbook_name = input()
print("\n")
print("Enter the exact name of the sheet within this workbook that should be converted:")
excel_sheet_name = input()




data_file = ("./data/input/"+workbook_name).strip()
# choose sheet
sheet_name= excel_sheet_name.strip()
# read excel file
df = pd.read_excel(data_file,
                   sheet_name=sheet_name,
                   index_col=[0],
                   na_values=['b.d.'])




# take description from top left cell
description = df.index.names[0]




# use "SAMPLE NAME" as index header
df.index = df.index.set_names(df.index[3])




# create mask to be able to distinguish difference in the column structure
mask = df.iloc[3,].isnull().values




# gather column names from two different rows
column_names = list(df.iloc[3,~mask].values) + list(df.iloc[0,mask].values)




column_names = [s.replace('/', ' ') for s in column_names]




df.columns = column_names




df.loc["flag"] = ["1.{}".format(i+1) if j == False else "2.{}".format(i+11) for i,j in enumerate(mask)]




# convert dataframe to xarray
xr = df.iloc[5:,].to_xarray()
# global attributes
xr.attrs = {'Conventions': 'CF-1.6', 'Title': sheet_name, 'Description': description}

# add variable attributes
units = df.iloc[2,mask]
method_codes = df.iloc[1,mask]




for i, col in enumerate(df.columns[mask]):
    getattr(xr, col).attrs = {'units': units[i], 'comment': 'METHOD CODE: {}'.format(method_codes[i])}




comments = df.iloc[4,~mask]

for i, col in enumerate(df.columns[~mask]):
    getattr(xr, col).attrs = {'comment': comments[i]}




# write xarray to netcdf file
xr.to_netcdf('./data/output/{}.nc'.format('-'.join((workbook_name.replace(".","")+" "+excel_sheet_name).split())))






