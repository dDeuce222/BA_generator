import shutil
import pandas as pd
import os

def save(file_name , bstring ,bcstring,lstring ,lcstring, blstring, count,bfstring,lfstring):
    if not os.path.exists('Back'):
        os.makedirs('Back')
    fp = open('Back/' + file_name + '_back.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + bstring
    fp.write(context)
    fp.close()
    if not os.path.exists('BackFlat'):
        os.makedirs('BackFlat')
    fp = open('BackFlat/' + file_name + '_backflat.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + bfstring
    fp.write(context)
    fp.close()
    if not os.path.exists('BackCond'):
        os.makedirs('BackCond')
    fp = open('BackCond/' + file_name + '_backcond.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + bcstring
    fp.write(context)
    fp.close()
    if not os.path.exists('Lay'):
        os.makedirs('Lay')
    fp = open('Lay/' + file_name + '_lay.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + lstring
    fp.write(context)
    fp.close()
    if not os.path.exists('LayFlat'):
        os.makedirs('LayFlat')
    fp = open('LayFlat/' + file_name + '_layflat.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + lfstring
    fp.write(context)
    fp.close()
    if not os.path.exists('LayCond'):
        os.makedirs('LayCond')
    fp = open('LayCond/' + file_name + '_laycond.baf','w')
    context = "4.0\n\n\n" + str(count)
    context = context + lcstring
    fp.write(context)
    fp.close()
    if not os.path.exists('Back&Lay'):
        os.makedirs('Back&Lay')
    fp = open('Back&Lay/' + file_name + '_backlay.baf','w')
    context = "4.0\n\n\n" + str(count*2)
    context = context + blstring
    fp.write(context)
    fp.close()


if(os.path.exists('Back')):
    shutil.rmtree('Back')
if(os.path.exists('BackCond')):
    shutil.rmtree('BackCond')
if(os.path.exists('Lay')):
    shutil.rmtree('Lay')
if(os.path.exists('LayCond')):
    shutil.rmtree('LayCond')
if(os.path.exists('Back&Lay')):
    shutil.rmtree('Back&Lay')
loc = input("Please insert refromatted excel file location : ")

#read xlsx file into panda dataframe 
df = pd.read_excel(loc,engine="openpyxl")

#get column names
index = df.columns
rule_file_name = ''
for i in range(0,len(df)):
    #get row data
    row = df.loc[i]
    #get COL_A of row i , strip to remove whitespace , lower to enable matching in both lower and upper case
    
    name = row['Time']
    selection = row['Selection']
    value = row['Value']
    target = row['Target']
    print(i,name , rule_file_name)
    if(name != rule_file_name):
        if(rule_file_name != ''):
            save(str(int(rule_file_name)),back_save_string,back_cond_string,lay_save_string,lay_cond_string,test_save_string,j , back_flat_string,lay_flat_string)
        rule_file_name = name
        j = 0 # if j bigger than 9 , needs format changing
        back_save_string = ''
        back_cond_string = ''
        back_flat_string = ''
        lay_cond_string = ''
        lay_save_string = ''
        lay_flat_string = ''
        test_save_string = ''
    back_string = """
000{No}
{No}
BET_{Type}
1
False
True
30
False
True
0
5
1
2
{No}

SECOND_BEST

1
{Back}
{STAKE}
False
False
False
False
False



False


1
SELECTION_COUNT
Number of selections = {Count}
2
EQUAL
{Count}
"""
    test_string = """
000{No_1}
{No_2} {Type}
BET_{Type}
1
False
True
30
False
True
0
5
1
2
{No_2}

SECOND_BEST

1
{Value}
LIABILITY
False
False
False
False
False



False


2
SELECTION_COUNT
Number of selections = {Count}
2
EQUAL
{Count}
FIXED_ODDS
{Type_1} price {BigSmall} {Target}
7
{Target}
{Compare}
{Type}
True
2
1

"""
    
    j += 1
    selection = int(selection)
    back_save_string += back_string.format(No = selection ,Back = round(value,2) , Count = 6,Type='BACK',STAKE='LIABILITY')
    back_flat_string += back_string.format(No = selection ,Back = round(value,2) , Count = 6,Type='BACK',STAKE='FIXED')
    lay_save_string += back_string.format(No = selection ,Back = round(value,2) , Count = 6,Type='LAY',STAKE='LIABILITY')
    lay_flat_string += back_string.format(No = selection ,Back = round(value,2) , Count = 6,Type='LAY', STAKE='FIXED')
    back_cond_string += test_string.format(No_1 = selection,No_2 = selection,Type = 'BACK' , Value = round(value,2),Count = 6 , Type_1 = 'Back',Target = round(target,2), BigSmall = '>',Compare='GREATER')
    lay_cond_string += test_string.format(No_1 = selection,No_2 = selection,Type = 'LAY' , Value = round(value,2),Count = 6 , Type_1 = 'Lay',Target = round(target,2), BigSmall = '<',Compare='LESS')
    test_save_string += test_string.format(No_1 = str(selection*2-1),No_2 = selection,Type = 'BACK' , Value = round(value,2),Count = 6 , Type_1 = 'Back',Target = round(target,2), BigSmall = '>',Compare='GREATER')
    test_save_string += test_string.format(No_1 = str(selection*2),No_2 = selection,Type = 'LAY' , Value = round(value,2),Count = 6 , Type_1 = 'Lay',Target = round(target,2), BigSmall = '<',Compare='LESS')
save(str(int(rule_file_name)),back_save_string,back_cond_string,lay_save_string,lay_cond_string,test_save_string,j,back_flat_string,lay_flat_string)    

    
    
    



#saving file
# save_path = 'Results'
# name_of_file = input("What do you want to save the file as?   ")
# completeName = os.path.join(save_path, name_of_file+".xlsx")         
# #mkdir if not exists
# if not os.path.exists(save_path):
#     os.makedirs(save_path)

# #convert to dataframe
# result_df = pd.DataFrame({'WORD':keywords,'COUNTS':word_counts,'IMPUTATION':imputations})
# # Create a Pandas Excel writer using XlsxWriter as the engine.
# writer = pd.ExcelWriter(completeName, engine='xlsxwriter')

# # Convert the dataframe to an XlsxWriter Excel object.
# result_df.to_excel(writer, sheet_name='Sheet1', index=False)

# # Close the Pandas Excel writer and output the Excel file.
# writer.save()"""