######### SCRIPT Number 3 #############


import pandas as pd


df = pd.read_excel('final.xlsx')

for row_index, row in df.iterrows():
    #print(row[0], row[1])
    index = row_index
    if row[0] == '§§':
        #print(index)
        word, label = df.iloc[[index][0]]
        #print(index,word, label)
        word = " "
        label = " "

        df.iloc[[index][0]] = word, label

print('replaced §§ tags with empty lines')

for row_index, row in df.iterrows():
    print(row[0], row[1])

df.to_excel('final_train_data.xlsx')