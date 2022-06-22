import openpyxl as ox
import os
import csv
import pandas
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
# # read_file=pd.read_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.csv')
# # read_file.to_excel(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.xlsx')
# print(df)
#
# #skip rows =1 will not read rows you ask or header=1 row
# add header header=None, names["ex1", "ex2", "ex3"]

########################################################################################################################

bw7=pd.read_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\7-BW.csv')
bw10=pd.read_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\10-BW.csv')
bw11=pd.read_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\11-BW.csv')
frame = [bw7, bw10, bw11]
bw_master = pd.concat(frame)

bw_master.to_csv(r'C:\Users\Joshua Kemperman\OneDrive - Enchante Living\desktop enchliving\WH File\BW_Master.csv', index = False)

