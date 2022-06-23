#import openpyxl as ox
#import os
import csv
#import pandas
import pandas as pd
import glob


# df=pd.read_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.csv')
# # read_file.to_excel(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.xlsx')
# #
# #
# #
# # book=ox.load_workbook(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.xlsx')
# # sheet =book["Sheet1"]
# #
# # print(sheet["E2"].value)
# #
# s=os.getcwd()
# print(s)
# p=os.chdir(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File')
# # print(os.getcwd())
# # for xls in os.listdir():
# #     if xls.startswith("7" or "11" or "13"):
# #        print(xls)
# #
# #
# # list=[7-BW.csv, 7-DR.csv,7-HE.csv, 7-ON.csv, 7-S1.csv, 7-SI.csv, 7-SV.csv]
# #
# #
# # read_file=pd.read_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\7-BW.csv')
# # read_file.to_excel(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\7-BW.xlsx')
# print(df)
#
# #skip rows =1 will not read rows you ask or header=1 row
# add header header=None, names["ex1", "ex2", "ex3"]


############################################### PATHS ##################################################################
si7_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-SI.csv'
si13_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-SI.csv'
on7_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-ON.csv'
on13_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-ON.csv'
sv7_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-SV.csv'
sv13_p = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\13-SV.csv'

save_Loc = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project'

############################################CREATE DATA FRAMES##########################################################

# bw7=pd.read_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\7-BW.csv')
# bw10=pd.read_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\10-BW.csv')
# bw11=pd.read_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Raw Files\11-BW.csv')
# frame = [bw7, bw10, bw11]
# bw_master = pd.concat(frame)


si7=pd.read_csv(si7_p)
si13=pd.read_csv(si13_p)
frame = [si7, si13]
sisales_master = pd.concat(frame)

on7=pd.read_csv(on7_p)
on13=pd.read_csv(on13_p)
frame=[on7, on13]
on_sales_master = pd.concat(frame)

sv7=pd.read_csv(sv7_p)
sv13=pd.read_csv(sv13_p)
frame = [sv7, sv13]
sv_sales_master = pd.concat(frame)



################################################ FILE OUTPUTS ##########################################################

#bw_master.to_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\BW_Master.csv', index = False)
#si_master.to_csv(r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\SI_Master.csv', index = False)

fileName = pd.ExcelWriter(save_Loc, engine = 'xlsxwriter')

# bw_master.to_excel(fileName, sheet_name='BW', index = False)
si_master.to_excel(fileName, sheet_name='SI Sales', index = False)
on_master.to_excel(fileName, sheet_name='ON Sales', index = False)
sv_master.to_excel(fileName, sheet_name='SV Sales', index = False)



fileName.save()
