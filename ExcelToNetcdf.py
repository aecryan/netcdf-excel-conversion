#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!/usr/bin/env python3
import pandas as pd


# In[2]:


#sheet names:
# 3 WR Major and trace elements
# 4 Amphibole Major elements
# 5 OPX Major elements
# 6 Glass Major elements
# 7 Plagioclase Major elements
# 8 Plagioclase trace elements


# In[3]:


data_file = '../data/Elemental_BulkSample_EarthChem_Example_1395-1_Weber_Nevado_del_Toluca (1).xls'
# choose sheet
sheet_name='3 WR Major and trace elements'
# read excel file
df = pd.read_excel(data_file,
                   sheet_name=sheet_name,
                   index_col=[0],
                   na_values=['b.d.'])


# In[4]:


# take description from top left cell
description = df.index.names[0]


# In[5]:


# use "SAMPLE NAME" as index header
df.index = df.index.set_names(df.index[3])


# In[6]:


# create mask to be able to distinguish difference in the column structure
mask = df.iloc[3,].isnull().values


# In[7]:


# gather column names from two different rows
column_names = list(df.iloc[3,~mask].values) + list(df.iloc[0,mask].values)


# In[8]:


column_names = [s.replace('/', ' ') for s in column_names]


# In[9]:


df.columns = column_names


# In[10]:


# df.loc["flag"] = ["1" if j == False else "2" for j in mask]


# In[11]:


# This is added by Amir just for internal communication
# It preserves the order of the columns and also the left and right side of the table
# It is very strange that the following version of the code does not work: 
# df.loc["flag"] = ["left-{}".format(i) if j == False else "right-{}".format(i) for i,j in enumerate(mask)]
df.loc["flag"] = ["1.{}".format(i+1) if j == False else "2.{}".format(i+11) for i,j in enumerate(mask)]


# In[12]:


# convert dataframe to xarray
xr = df.iloc[5:,].to_xarray()
# global attributes
xr.attrs = {'Conventions': 'CF-1.6', 'Title': sheet_name, 'Description': description}

# add variable attributes
units = df.iloc[2,mask]
method_codes = df.iloc[1,mask]


# In[13]:


for i, col in enumerate(df.columns[mask]):
    getattr(xr, col).attrs = {'units': units[i], 'comment': 'METHOD CODE: {}'.format(method_codes[i])}
#     print(getattr(xr, col).attrs)
#     print("--------------------------------")


# In[14]:


comments = df.iloc[4,~mask]

for i, col in enumerate(df.columns[~mask]):
    getattr(xr, col).attrs = {'comment': comments[i]}


# In[15]:


# write xarray to netcdf file
xr.to_netcdf('../data/{}.nc'.format('-'.join(sheet_name.split())))


# In[16]:


xr

