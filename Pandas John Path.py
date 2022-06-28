# Authors: John Ayres, Joshua Kemperman

import pandas as pd


######################################################## PATHS ########################################################
si7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI.csv'
si13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI.csv'
on7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON.csv'
on13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON.csv'
sv7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV.csv'
sv13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV.csv'

si7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI-Inv.xls'
si13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI-Inv.xls'
on7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON-Inv.xls'
on13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON-Inv.xls'
sv7_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV-Inv.xls'
sv13_inv_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV-Inv.xls'

iv_rep = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\05 - INVENTORY\Inventory Report - For Audit - June 20, 2022.xlsx'

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

si7_inv=pd.read_excel(si7_inv_p,engine= 'xlrd')
si13_inv=pd.read_excel(si13_inv_p,engine= 'xlrd')
on7_inv=pd.read_excel(on7_inv_p,engine= 'xlrd')
on13_inv=pd.read_excel(on13_inv_p,engine= 'xlrd')
sv7_inv=pd.read_excel(sv7_inv_p,engine= 'xlrd')
sv13_inv=pd.read_excel(sv13_inv_p,engine= 'xlrd')

frame = [si7_inv, si13_inv]
si_inv_master = pd.concat(frame)

frame = [on7_inv, on13_inv]
on_inv_master = pd.concat(frame)

frame = [sv7_inv, sv13_inv]
sv_inv_master = pd.concat(frame)

inv_rep =pd.read_excel(iv_rep,header=2)

##################################################### Changes ##########################################################

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

si_inv_master["WH"] = 'SI'
on_inv_master["WH"] = 'ON'
sv_inv_master["WH"] = 'SV'

# Rename units case cubic
si_inv_master['unit']=si_inv_master['si']
del si_inv_master['si']
on_inv_master['unit']=on_inv_master['on']
del on_inv_master['on']
sv_inv_master['unit']=sv_inv_master['sv']
del sv_inv_master['sv']

si_inv_master['caseqty']=si_inv_master['si_caseqty']
del si_inv_master['si_caseqty']
on_inv_master['caseqty']=on_inv_master['on_caseqty']
del on_inv_master['on_caseqty']
sv_inv_master['caseqty']=sv_inv_master['sv_caseqty']
del sv_inv_master['sv_caseqty']

si_inv_master['cubic']=si_inv_master['si_cubic']
del si_inv_master['si_cubic']
on_inv_master['cubic']=on_inv_master['on_cubic']
del on_inv_master['on_cubic']
sv_inv_master['cubic']=sv_inv_master['sv_cubic']
del sv_inv_master['sv_cubic']


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

si_inv_master = si_inv_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]
on_inv_master = on_inv_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]
sv_inv_master = sv_inv_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]

#data type manip
si_sales_master['WH'] = si_sales_master['WH'].astype(str)
si_sales_master['Month'] = si_sales_master['Month'].astype(float)
si_sales_master['amount'] = si_sales_master['amount'].astype(float)

################################################### Creating Master sheets #############################################
frame = [si_sales_master,on_sales_master,sv_sales_master]
sales_master = pd.concat(frame)

frame = [si_inv_master,on_inv_master,sv_inv_master]
inv_master = pd.concat(frame)


####################################################### Changes ########################################################
# VLOOKUP SIM
sales_master["ID"] = sales_master["style"]+sales_master["color"]

inv_rep = inv_rep[['ID','CRTN CBM','Casepack']]

sales_master = pd.merge(sales_master,inv_rep,on='ID',how='inner')

inv_master = inv_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]

# Rounding
sales_master['Total CBM']=round((sales_master['CRTN CBM']/sales_master['Casepack'])*sales_master['shptot'],4)
sales_master['CRTN CBM'] =round(sales_master['CRTN CBM'],4)

#### Summary sheets
#Sales by wh by month

sales_month = pd.DataFrame()
sales_month_master = pd.DataFrame()

sales_month = sales_master[sales_master.WH =='ON']
sales_month = sales_month[['WH','Month','amount']]
sales_month_on_01 =sales_month[sales_month.Month =='01']
sales_month_master[1,0] = sum(sales_month_on_01['amount'])
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

# SUMMARY OUT
sales_master.to_excel(fileName, sheet_name='Total Sales', index = False)
inv_master.to_excel(fileName, sheet_name='Total INV', index = False)

#sales_month.to_excel(fileName, sheet_name='Sales Table', index = False)
#sales_month_on_01.to_excel(fileName, sheet_name='Sales on01', index = False)
#sales_month_master.to_excel(fileName, sheet_name='Sales master', index = False)

fileName.save()
