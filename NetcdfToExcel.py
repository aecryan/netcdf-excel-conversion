#!/usr/bin/env python
# coding: utf-8

#Project name: netCDF --> Excel
#Description: It receives a netCDF file with a predefined format and returns the equivalent file in excel format. 
#Programmers: Amir \& Ashley
#Date: 07-09-2020




#Loading the packages
import pandas as pd
import xarray as xr
import numpy as np




print("Enter the exact name of the netCDF file. Please do include the file extension (e.g., .nc):")
netCDF_file_name = input()





ds_disk = xr.open_dataset("./data/input/"+netCDF_file_name)




#Fetching the global attributes
ds_global_attributes = ds_disk.attrs




#Fetching the indexes (It will not be used later on as the indeces are also available in the variables)
ds_indeces = list(ds_disk.coords.indexes["SAMPLE NAME"])




#Fetching the variables
ds_disk_variables = ds_disk.to_dataframe().reset_index()
#Fetching the column names
column_names = list(ds_disk_variables.columns)




# Retrieve the columns order
orders = ds_disk_variables.iloc[len(ds_disk_variables)-1].to_dict()

# Remove the columns order
ds_disk_variables = ds_disk_variables.drop(axis=0, index=len(ds_disk_variables)-1)
# ds_disk_variables = ds_disk_variables.drop(axis=0, index=len(ds_disk_variables))

# Extraction of the order and place them in two arrays
column_dict = {"left":{}, "right":{}}
for i,j in orders.items():
    info = str(j).split(".")
    if j == "flag":
        side, order = 1 , 0
        column_dict["left"][i] = "1.0"
    else:
        side, order = info[0] , info[1]
        if side == "1":
            column_dict["left"][i] = j
        elif side == "2":
            column_dict["right"][i] = j

left_dict = {k: v for k, v in sorted(column_dict["left"].items(), key=lambda item: item[1])}
right_dict = {k: v for k, v in sorted(column_dict["right"].items(), key=lambda item: item[1])}

left = [i for i in left_dict]
right = [i for i in right_dict]

# Reordering the columns
ds_disk_variables = ds_disk_variables[left+right]




#Creating an empty dictionary where the keys are the column names, and the values refer to the properties of those columns 

ds_dict = {i:[] for i in left+right}
for j in left:
    ds_dict[j] = {}
    ds_dict[j]["comment"] = ""

for j in right:
    ds_dict[j] = {}
    ds_dict[j]["units"] = ""
    ds_dict[j]["comment"] = ""




#Populating the dictionary 
for i in left:
    if i == "SAMPLE NAME":
        #Is this property already stored in the NetCDF file?
        ds_dict[i]["comment"] = "must match a sample on the SAMPLES tab column A"
    else:
        ds_dict[i]["comment"] = ds_disk.data_vars[i].attrs["comment"]

        
for j in right:
    ds_dict[j]["units"] = ds_disk.data_vars[j].attrs["units"]
    ds_dict[j]["comment"] = ds_disk.data_vars[j].attrs["comment"]




#Creating a dataframe for the left side of the tables's header
s = pd.DataFrame(columns=ds_disk_variables.columns)
s.loc[0] = [ds_dict[i]["comment"] if i in left else np.nan for i in s.columns] 

g = pd.DataFrame(columns=ds_disk_variables.columns)
g.loc[0] = [i if i in left else np.nan for i in ds_disk_variables.columns]

a1 = pd.concat([g,s])




#Creating a dataframe for the right side of the tables's header

h = pd.DataFrame(columns=ds_disk_variables.columns)
h.loc[0] = [ds_dict[i]["comment"].replace("METHOD CODE: ","") if i in right else np.nan for i in s.columns]
h.loc[1] = [ds_dict[i]["units"] if i in right else np.nan for i in s.columns]

v = pd.DataFrame(columns=ds_disk_variables.columns)
v.loc[0] = [i if i in right else np.nan for i in ds_disk_variables.columns]

a2 = pd.concat([v,h])




#Manual insertion of some values in the table (Do they exist in the NetCDF file? if so, where?)
a2.iloc[0]["number of replicates"] = "PARAMETER [list]"
a2.iloc[1]["number of replicates"] = "METHOD CODE [more info]:"
a2.iloc[2]["number of replicates"] = "UNIT [list]:"




#Here, the header to the table is created
a3 = pd.concat([a2,a1])




#Here, the header and variables are added together
a4 = pd.concat([a3,ds_disk_variables])




#Changing the columns names
k = [np.nan for i in range(len(left+right))]
k[0] = ds_global_attributes["Description"]

a4.columns = k




#Replacing the dataframe index by the column "SAMPLE NAME"
a4.index = a4[ds_global_attributes["Description"]]

#Removing the column "SAMPLE NAME" from the dataframe
a4 = a4.drop(columns=[ds_global_attributes["Description"]])




#Saving the dataframe as an excel file
a4.to_excel("./data/output/"+netCDF_file_name.replace(".","")+".xlsx", sheet_name=ds_global_attributes["Title"])

