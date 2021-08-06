import pandas as pd

df1 = pd.read_excel('classes_excel/classes_lower_case.xlsx')
#print(df1)

MOD =  []
SPEC = []
MAT =  []
SEC =  []
MED =  []
PROC = []
ORG =  []
NORM = []


########################################################################################################################

for row_index, row in df1.iterrows():
    MOD.append(row[0])
    SPEC.append(row[1])
    MAT.append(row[2])
    SEC.append(row[3])
    MED.append(row[4])
    PROC.append(row[5])
    ORG.append(row[6])
    NORM.append(row[7])


"""
print(MOD)
print(SPEC)
print(MAT)
print(SEC)
print(MED)
print(PROC)
print(ORG)
print(NORM)
"""

################ to remove nan values #################################################################################

mod_newlist = [x for x in MOD if pd.isnull(x) == False]
spec_newlist = [x for x in SPEC if pd.isnull(x) == False]
mat_newlist = [x for x in MAT if pd.isnull(x) == False]
sec_newlist = [x for x in SEC if pd.isnull(x) == False]
med_newlist = [x for x in MED if pd.isnull(x) == False]
proc_newlist = [x for x in PROC if pd.isnull(x) == False]
org_newlist = [x for x in ORG if pd.isnull(x) == False]
norm_newlist = [x for x in NORM if pd.isnull(x) == False]

print('module list:', mod_newlist)
print('specification list:', spec_newlist)
print('material list:', mat_newlist)
print('security list:', sec_newlist)
print('medien list:', med_newlist)
print('processes list:', proc_newlist)
print('organizatons list', org_newlist)
print('standards list:', norm_newlist)

######################################################################################################################
#####################################################################################################################


df2 = pd.read_excel('cleaned.xlsx')
#print(df2)

for row_index, row in df2.iterrows():
    current_index = row_index
    prev_index = current_index - 1
    prev_prev_index = current_index - 2

    pword, plabel = df2.iloc[[prev_index][0]]        # prev word from current index
    ppword, pplabel = df2.iloc[[prev_prev_index][0]] # prev to prev word from current index

    ########################## check for MODULES #######################################################
    if row[0] in mod_newlist:
        # print('found MOD')
        row[1] = 'MOD'
        index = row_index

        word, label = df2.iloc[[index][0]]
        label = 'MOD'
        df2.iloc[[index][0]] = word, label
    ######################################################################################################

    ####################### check for Specifications ####################################################
    elif row[0] in spec_newlist:
        # print('found SPEC')
        row[1] = 'SPEC'

        index = row_index
        word, label = df2.iloc[[index][0]]
        label = 'SPEC'
        df2.iloc[[index][0]] = word, label
    ######################################################################################################

    #################### check for MATERIALS #############################################################
    elif row[0] in mat_newlist:
        # print('found MAT')
        row[1] = 'MAT'

        index = row_index
        word, label = df2.iloc[[index][0]]
        label = 'MAT'
        df2.iloc[[index][0]] = word, label
    ######################################################################################################

    #################ä check for Security ################################################################
    elif row[0] in sec_newlist:
        # print('found SEC')
        row[1] = 'SEC'
        index = row_index
        word, label = df2.iloc[[index][0]]
        label = 'SEC'
        df2.iloc[[index][0]] = word, label
    ####################################################################################################

    #################### check for Medien ##############################################################
    elif row[0] in med_newlist:
        # print('found MED')
        row[1] = 'MED'

        index = row_index
        word, label = df2.iloc[[index][0]]
        label = 'MED'
        df2.iloc[[index][0]] = word, label
    #####################################################################################################

    #################### check for Processes ############################################################
    elif row[0] in proc_newlist:
        # print('found PROC')
        row[1] = 'PROC'

        index = row_index
        word, label = df2.iloc[[index][0]]
        label = 'PROC'
        df2.iloc[[index][0]] = word, label
    #####################################################################################################

    #################  check for Organizations #########################################################

    elif row[0] in org_newlist:

        #print('found B-ORG')
        #row[1] = 'B-ORG'
        current_index = row_index
        next_index = current_index + 1
        next_next_index = current_index + 2

        cword, clabel = df2.iloc[[current_index][0]]
        nword, nlabel = df2.iloc[[next_index][0]]  # next word from current index
        nnword, nnlabel = df2.iloc[[next_next_index][0]]

        #print('C', clabel)
        #print('N', nlabel)
        #print('NN', nnlabel)

        if clabel not in ['B-ORG', 'I-ORG', 'ORG']:
            #print('we are labelling')

            # eg woll gmbh
            if nword == 'gmbh':
                #print('in if')
                label1 = 'B-ORG'
                label2 = 'I-ORG'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[next_index][0]] = nword, label2

            elif nword == '&' and nnword in org_newlist:
                #print('in elif')
                label1 = 'B-ORG'
                label2 = 'I-ORG'
                label3 = 'I-ORG'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[next_index][0]] = nword, label2
                df2.iloc[[next_next_index][0]] = nnword, nnlabel

            elif cword == 'gmbh':
                label1 = 'I-ORG'
                df2.iloc[[current_index][0]] = cword, label1


            else:
                label1 = 'B-ORG'
                df2.iloc[[current_index][0]] = cword, label1

    ####################################################################################################################

    ################### check for NORMS / Standards ###################################################################

    elif row[0] in norm_newlist :
         #print('found din')

         current_index = row_index
         index2 = current_index + 1
         index3 = current_index + 2
         index4 = current_index + 3
         index5 = current_index + 4
         index6 = current_index + 5
         index7 = current_index + 6

         cword, clabel = df2.iloc[[current_index][0]]
         word2, label2 = df2.iloc[[index2][0]]
         word3, label3 = df2.iloc[[index3][0]]
         word4, label4 = df2.iloc[[index4][0]]
         word5, label5 = df2.iloc[[index5][0]]
         word6, label6 = df2.iloc[[index6][0]]
         word7, label7 = df2.iloc[[index7][0]]

         #print("cword:", cword)
         #print("word2:", word2)
         #print("word3:", word3)
         #print("word4:", word4, type(word4))
         #print("word5:", word5)
         #print("word6:", word6)


         if clabel not in ['B-NORM', 'I-NORM', 'NORM']:
            #print('we are labelling')

            # Case1a : din en iso number : number
            if cword == "din" and word2 == "en" and word3 == "iso" and word4.isnumeric() and word5 == ":" and word6.isnumeric():
                print('in case1a')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2
                df2.iloc[[index6][0]] = word6, label2

            # Case1b : din en number - number
            elif cword == "din" and word2 == "en" and word3.isnumeric() and word4 == "-" and word5.isnumeric():
                print('in case 1b')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2

            # Case 1c: din en number
            elif cword == "din" and word2 == "en" and word3.isnumeric():
                print("in case 1c")
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2

            # Case 1d:  din - vde number - number
            elif cword == "din" and word2 == "-" and word3 == "vde" and word4.isnumeric() and word5 == "-" and word6.isnumeric():
                print("in case 1d")
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2
                df2.iloc[[index6][0]] = word6, label2

            # Case 1e: din - vde number

            elif cword == "din" and word2 == "-" and word3 == "vde" and word4.isnumeric():
                print('in case 1e')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2

            # Case 2a: en iso number - number : number
            elif cword == "en" and word2 == "iso" and word3.isnumeric() and word4 == "-" and word5.isnumeric() and word6 == ":" and word7.isnumeric():
                print('in case 2a')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2
                df2.iloc[[index6][0]] = word6, label2
                df2.iloc[[index7][0]] = word7, label2

            # Case2b: en iso number : number
            elif cword == "en" and word2 == "en" and word3.isnumeric() and word4 == ":" and word5.isnumeric():
                print('in case 2b')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2

            # Case 2c: en number : number OR en number - number
            elif cword == "en" and word2.isnumeric() and (word3 == ":" or word3 == "-") and word4.isnumeric():
                print('in case 2c')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2

            # Case 2d: en iso number
            elif cword == "en" and word2 == "iso" and word3.isnumeric():
                print('in case 2d')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2

            # Case 3a : iso number - number : number
            elif cword == "iso" and word2.isnumeric() and word3 == "-" and word4.isnumeric() and word5 == ":" and word6.isnumeric():
                print("in case 3a")
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2
                df2.iloc[[index5][0]] = word5, label2
                df2.iloc[[index6][0]] = word6, label2

            # Case 3b : iso number : number
            elif cword == "iso" and word2.isnumeric() and word3 == ":" and word4.isnumeric():
                print(' in case 3b')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2
                df2.iloc[[index3][0]] = word3, label2
                df2.iloc[[index4][0]] = word4, label2

            # Case 3c: vdi number
            elif cword == "vdi" and word2.isnumeric():
                print('in case 3c')
                label1 = 'B-NORM'
                label2 = 'I-NORM'

                df2.iloc[[current_index][0]] = cword, label1
                df2.iloc[[index2][0]] = word2, label2

            elif row[0] in norm_newlist:
                label1 = 'B-NORM'
                df2.iloc[[current_index][0]] = cword, label1
                
         else:
            continue

########################################################################################################################




#print(row[0], row[1])

########################################################################################################################


df2.to_excel('final.xlsx')
print('saved to excel')