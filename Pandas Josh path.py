#import openpyxl as ox
#import os
import csv
#import pandas
import pandas as pd
import glob
import datetime




######################################################## PATHS ########################################################
si7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI.csv'
si13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI.csv'
on7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON.csv'
on13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON.csv'
sv7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV.csv'
sv13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV.csv'

#### CHANGE DATE TO TODAY
save_Loc = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\US, CA WHSE Sales and OH 06.23.22.xlsx'




################################################## CREATE DATA FRAMES ##################################################

si7=pd.read_csv(si7_p)
si13=pd.read_csv(si13_p)
frame = [si7, si13]
si_sales_master = pd.concat(frame)

on7=pd.read_csv(on7_p)
on13=pd.read_csv(on13_p)
frame=[on7, on13]
on_sales_master = pd.concat(frame)

sv7=pd.read_csv(sv7_p)
sv13=pd.read_csv(sv13_p)
frame = [sv7, sv13]
sv_sales_master = pd.concat(frame)

si7_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI-Inv.xls')
si13_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI-Inv.xls')
on7_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON-Inv.xls')
on13_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON-Inv.xls')
sv7_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV-Inv.xls')
sv13_inv=pd.read_excel(r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV-Inv.xls')

frame = [si7_inv, si13_inv]
si_inv_master = pd.concat(frame)

frame = [on7_inv, on13_inv]
on_inv_master = pd.concat(frame)

frame = [sv7_inv, sv13_inv]
sv_inv_master = pd.concat(frame)


############################################### Changes ####################################################

# Creating Month Col
si_sales_master['invdate'] = pd.to_datetime(si_sales_master['invdate'])
si_sales_master['Month'] = si_sales_master['invdate'].dt.strftime('%m')
on_sales_master['invdate'] = pd.to_datetime(on_sales_master['invdate'])
on_sales_master['Month'] = on_sales_master['invdate'].dt.strftime('%m')
sv_sales_master['invdate'] = pd.to_datetime(sv_sales_master['invdate'])
sv_sales_master['Month'] = sv_sales_master['invdate'].dt.strftime('%m')

# Creating WH col
si_sales_master["WH"] = 'SI'
on_sales_master["WH"] = 'ON'
sv_sales_master["WH"] = 'SV'

# Col Reorg
si_sales_master = si_sales_master[['WH','Month','division','linecode','style','color','desc','catcode','upcno','shptot',
                                   'price','amount','cost','account','name','custpo','invoice','invdate','order',
                                   'entered','edifile','note1','ststate','stzip','piktkt','piktktdate']]
on_sales_master = on_sales_master[['WH','Month','division','linecode','style','color','desc','catcode','upcno','shptot',
                                   'price','amount','cost','account','name','custpo','invoice','invdate','order',
                                   'entered','edifile','note1','ststate','stzip','piktkt','piktktdate']]
sv_sales_master = sv_sales_master[['WH','Month','division','linecode','style','color','desc','catcode','upcno','shptot',
                                   'price','amount','cost','account','name','custpo','invoice','invdate','order',
                                   'entered','edifile','note1','ststate','stzip','piktkt','piktktdate']]

si_inv_master = si_inv_master[['style','grp','desc','color','division','cubic_ft','weight','master_pack','si',
                               'si_caseqty', 'si_cubic']]
on_inv_master = on_inv_master[['style','grp','desc','color','division','cubic_ft','weight','master_pack','on',
                               'on_caseqty', 'on_cubic']]
sv_inv_master = sv_inv_master[['style','grp','desc','color','division','cubic_ft','weight','master_pack','sv',
                               'sv_caseqty', 'sv_cubic']]
##################################################### FILE OUTPUTS #####################################################


fileName = pd.ExcelWriter(save_Loc, engine = 'xlsxwriter')

# Sales OUT
si_sales_master.to_excel(fileName, sheet_name='SI Sales', index = False)
on_sales_master.to_excel(fileName, sheet_name='ON Sales', index = False)
sv_sales_master.to_excel(fileName, sheet_name='SV Sales', index = False)

#INV OUT
si_inv_master.to_excel(fileName, sheet_name='SI INV', index = False)
on_inv_master.to_excel(fileName, sheet_name='ON INV', index = False)
sv_inv_master.to_excel(fileName, sheet_name='SV INV', index = False)


fileName.save()
