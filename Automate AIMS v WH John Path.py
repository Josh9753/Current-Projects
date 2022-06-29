# first project automate report for AIMS Stock vs WH Stock
import pandas as pd



on7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON.csv'
on13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON.csv'
sv7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV.csv'
sv13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV.csv'


on7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON-Inv.xls'
on13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON-Inv.xls'
sv7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV-Inv.xls'
sv13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV-Inv.xls'

iv_rep = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\05 - INVENTORY\Inventory Report - For Audit - June 20, 2022.xlsx'

#####################################################################################################################

#Add tabs for WH reports

###############################################################################################################################


on7=pd.read_csv(on7_p)
on13=pd.read_csv(on13_p)
frame=[on7, on13]
on_sales_master = pd.concat(frame)

sv7=pd.read_csv(sv7_p)
sv13=pd.read_csv(sv13_p)
frame = [sv7, sv13]
sv_sales_master = pd.concat(frame)

on7_inv=pd.read_excel(on7_inv_p,engine= 'xlrd')
on13_inv=pd.read_excel(on13_inv_p,engine= 'xlrd')
sv7_inv=pd.read_excel(sv7_inv_p,engine= 'xlrd')
sv13_inv=pd.read_excel(sv13_inv_p,engine= 'xlrd')


frame = [on7_inv, on13_inv]
on_inv_master = pd.concat(frame)

frame = [sv7_inv, sv13_inv]
sv_inv_master = pd.concat(frame)

inv_rep =pd.read_excel(iv_rep,header=2)
