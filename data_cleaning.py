#### SCRIPT Number 1 ######


import numpy as np
import pandas as pd


### We read the excel file and remove empty lines from it ####################################

df2 = pd.read_excel('dataset/excel/tokenized_lower_case_train_data.xlsx')

ls = []

for row_index, row in df2.iterrows():
    if row[0] == " ": # removes empty lines
        ls.append(row_index)


# make new data frame
df3 = df2.drop(df2.index[ls])
#########################################################################################

# save the file
df3.to_excel('cleaned.xlsx')


