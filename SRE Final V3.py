#!/usr/bin/env python
# coding: utf-8

# # Final Project for ECEV 32000
# ## Jorge A. Alvarado

# In[1]:


import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import math
import matplotlib.pyplot as plt
import os
# Extracting data from SRE Excel output file
print("What is the file path of your raw data? For example: /Users/jorge/Documents/FILENAME.xlsx")
file_path = input()
firefly_sheet = pd.read_excel(file_path, sheet_name = 0)
renilla_sheet = pd.read_excel(file_path, sheet_name = 1)
firefly_sheet_array = np.array(firefly_sheet)
renilla_sheet_array = np.array(renilla_sheet)
# Slicing out raw data
firefly_raw_data = firefly_sheet_array[30:36, 3:12]
renilla_raw_data = renilla_sheet_array[30:36, 3:12]
firefly_over_renilla_raw_data = np.divide(firefly_raw_data, renilla_raw_data)
#Save raw data as csv.file
firefly_raw_data_df = pd.DataFrame(firefly_raw_data)
renilla_raw_data_df = pd.DataFrame(renilla_raw_data)
firefly_over_renilla_raw_data_df = pd.DataFrame(firefly_over_renilla_raw_data)
new_folder_data = "SRE Data Files"
if not os.path.exists(new_folder_data):
    os.makedirs(new_folder_data)
firefly_raw_data_df.to_csv("SRE Data Files/firefly_raw_data.csv", header=None, index=None)
renilla_raw_data_df.to_csv("SRE Data Files/renilla_raw_data.csv", header=None, index=None)
firefly_over_renilla_raw_data_df.to_csv("SRE Data Files/firefly_over_renilla_raw_data.csv", header=None, index=None)

#List all the sample names
print("Type the name of your samples, starting with your empty vector")
print('Type "done" when finished listing all samples')
total_sample_list = []
while True:
    sample_name = input()
    if str.lower(sample_name) == "done":
        break
    total_sample_list.append(sample_name)
print('\n')
#Assign your sample wells based on the coordinates of the displayed chart
print("Use this chart to assign coordinates (row,column) to the wells of your samples")
display(firefly_raw_data_df.fillna("empty"))
firefly_diction = {}
renilla_diction = {}
firefly_over_renilla_diction = {}
for i in total_sample_list:
    first_well_string = input("What are the coordinates for the first well of " + i + "? ")
    first_well_integer = first_well_string.split(',')
    first_well_integer = (int(first_well_integer[0]), int(first_well_integer[1]))
    last_well_string = input("What are the coordinates for the last well of " + i + "? ")
    last_well_integer = last_well_string.split(',')
    last_well_integer = (int(last_well_integer[0]), int(last_well_integer[1]))
    firefly_diction[i] = firefly_raw_data[first_well_integer[0]:last_well_integer[0] + 1, first_well_integer[1]:last_well_integer[1] + 1]
    renilla_diction[i] = renilla_raw_data[first_well_integer[0]:last_well_integer[0] + 1, first_well_integer[1]:last_well_integer[1] + 1]
    firefly_over_renilla_diction[i] =  firefly_over_renilla_raw_data[first_well_integer[0]:last_well_integer[0] + 1, first_well_integer[1]:last_well_integer[1] + 1]

#Store data to use for statistical analysis and plotting
#Firefly Stats
firefly_stats_list = []
for i in total_sample_list:
    stats_array = []
    stats_array.append(np.mean(firefly_diction[i]))
    stats_array.append(np.std(firefly_diction[i], ddof=1))
    stats_array.append(np.std(firefly_diction[i], ddof=1)/math.sqrt(firefly_diction[i].size))
    firefly_stats_list.append(stats_array)
firefly_stats_array = np.array(firefly_stats_list)
firefly_stats_df = pd.DataFrame(firefly_stats_array)
firefly_stats_df.iloc[0:len(total_sample_list)][0]
mean_firefly = list(firefly_stats_df.iloc[0:len(total_sample_list)][0])
error_firefly = list(firefly_stats_df.iloc[0:len(total_sample_list)][2])
#Renilla Stats
renilla_stats_list = []
for j in total_sample_list: 
    stats_array = []
    stats_array.append(np.mean(renilla_diction[j]))
    stats_array.append(np.std(renilla_diction[j], ddof=1))
    stats_array.append(np.std(renilla_diction[j], ddof=1)/math.sqrt(renilla_diction[j].size))
    renilla_stats_list.append(stats_array)
renilla_stats_array = np.array(renilla_stats_list)
renilla_stats_df = pd.DataFrame(renilla_stats_array)
renilla_stats_df.iloc[0:len(total_sample_list)][0]
mean_renilla = list(renilla_stats_df.iloc[0:len(total_sample_list)][0])
error_renilla = list(renilla_stats_df.iloc[0:len(total_sample_list)][2])
#Firefly/Renilla Stats
firefly_over_renilla_stats_list = []
for i in total_sample_list:
    stats_array = []
    stats_array.append(np.mean(firefly_over_renilla_diction[i]))
#To calculate STD of firefly over renilla, variance of the ratio must be calculated first using the covariance
    cov = np.cov(firefly_diction[i].astype(float), renilla_diction[i].astype(float))[0,1]
    varience = np.mean(firefly_diction[i])**2/np.mean(renilla_diction[i])**2 * (np.std(firefly_diction[i], ddof=1)**2/np.mean(firefly_diction[i])**2 - 2 * cov/(np.mean(firefly_diction[i]) * np.mean(renilla_diction[i])) + np.std(renilla_diction[i], ddof=1)**2/np.mean(renilla_diction[i])**2)
    stats_array.append(math.sqrt(varience))
    stats_array.append(math.sqrt(varience)/math.sqrt(firefly_diction[i].size))
    stats_array.append(varience)
    stats_array.append(cov)
    firefly_over_renilla_stats_list.append(stats_array)
firefly_over_renilla_stats_array = np.array(firefly_over_renilla_stats_list)
firefly_over_renilla_stats_df = pd.DataFrame(firefly_over_renilla_stats_array)
mean_firefly_over_renilla = list(firefly_over_renilla_stats_df.iloc[0:len(total_sample_list)][0])
error_firefly_over_renilla = list(firefly_over_renilla_stats_df.iloc[0:len(total_sample_list)][2])
#Fold Change Stats = Divide Firefly/Renilla by the mean of the EV
fold_change_stats_array = firefly_over_renilla_stats_array/firefly_over_renilla_stats_array[0,0]
fold_change_stats_df = pd.DataFrame(fold_change_stats_array)
mean_fold_change = list(fold_change_stats_df.iloc[0:len(total_sample_list)][0])
error_fold_change = list(fold_change_stats_df.iloc[0:len(total_sample_list)][2])


#Make 4 bar graphs that show each set of data along with statistics

#Construct Firefly Graph from data
x = total_sample_list
y = mean_firefly
fig, ax = plt.subplots()
ax.bar(x, y,  color ='red', width = 0.4, edgecolor = "black", yerr = error_firefly, align = 'center', capsize = 10)
ax.set_ylabel("Luminescence")
ax.set_xlabel("DNA Construct")
ax.set_title("FIREFLY")
plt.tight_layout()
plt.savefig('SRE Data Files/firefly.png')
plt.show()
#Show stats chart
firefly_stats_df.columns = ["Mean", "Standard Deviation", "Standard Error"]
firefly_stats_df.index = [total_sample_list]
display(firefly_stats_df)

#Construct Renilla Graph from data
x = total_sample_list
y = mean_renilla
fig, ax = plt.subplots()
ax.bar(x, y,  color ='green', width = 0.4, edgecolor = "black", yerr = error_renilla, align = 'center', capsize = 10)
ax.set_ylabel("Luminescence")
ax.set_xlabel("DNA Construct")
ax.set_title("RENILLA")
plt.tight_layout()
plt.savefig('SRE Data Files/renilla.png')
plt.show()
#Show stats chart
renilla_stats_df.columns = ["Mean", "Standard Deviation", "Standard Error"]
renilla_stats_df.index = [total_sample_list]
display(renilla_stats_df)

#Construct Firefly/Renilla Graph from data
x = total_sample_list
y = mean_firefly_over_renilla
fig, ax = plt.subplots()
ax.bar(x, y,  color ='yellow', width = 0.4, edgecolor = "black", yerr = error_firefly_over_renilla, align = 'center', capsize = 10)
ax.set_ylabel("Firefly/Renilla Ratio")
ax.set_xlabel("DNA Construct")
ax.set_title("FIREFLY/RENILLA")
plt.tight_layout()
plt.savefig('SRE Data Files/firefly_over_renilla.png')
plt.show()
#Show stats chart
firefly_over_renilla_stats_df.columns = ["Mean", "Standard Deviation", "Standard Error", "Varience", "Covarience"]
firefly_over_renilla_stats_df.index = [total_sample_list]
display(firefly_over_renilla_stats_df)

#Construct Fold Change Graph from data
x = total_sample_list
y = mean_fold_change
fig, ax = plt.subplots()
ax.bar(x, y,  color ='purple', width = 0.4, edgecolor = "black", yerr = error_fold_change, align = 'center', capsize = 10)
ax.set_ylabel("Fold Change")
ax.set_xlabel("DNA Construct")
ax.set_title("FOLD CHANGE")
plt.tight_layout()
plt.savefig('SRE Data Files/fold_change.png')
plt.show()
#Show stats chart
fold_change_stats_df.columns = ["Mean", "Standard Deviation", "Standard Error", "Varience", "Covarience"]
fold_change_stats_df.index = [total_sample_list]
display(fold_change_stats_df)

with pd.ExcelWriter('SRE Data Files/Stats.xlsx') as writer:
    firefly_stats_df.to_excel(writer, sheet_name='Firefly')
    renilla_stats_df.to_excel(writer, sheet_name='Renilla')
    firefly_over_renilla_stats_df.to_excel(writer, sheet_name='Firefly over Renilla')
    fold_change_stats_df.to_excel(writer, sheet_name='Fold Change')

print('You will find all output files in "SRE Data Files" folder at:' )
print(os.path.abspath("SRE Data File"))


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[19]:





# In[ ]:




