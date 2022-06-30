# Authors: John Ayres, Joshua Kemperman
# Desc: Script to process INV/Sales by WH, and Summary containing all INV/Sales JOHN PATH IN PROGRESS, ADDING FUNCTIONALITY
# PARAM: Input files as listed under pathing; Warehouse INV/OHS by data group and INV Report
# OUTPUT: Multipage excel Document, listing total INV/OHS by warehouse, Sheets containing totals of INV/OHS

import pandas as pd
import xlrd
import openpyxl
import xlsxwriter

######################################################## PATHS #########################################################
si7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SI.xls'
si13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SI.xls'
on7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-ON.xls'
on13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-ON.xls'
sv7_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\7-SV.xls'
sv13_p = r'C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\39 Joint Project\Raw Files\13-SV.xls'

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

si7 = pd.read_excel(si7_p,engine= 'xlrd')
si13 = pd.read_excel(si13_p,engine= 'xlrd')
frame = [si7, si13]
si_sales_master = pd.concat(frame,ignore_index=True)

on7 = pd.read_excel(on7_p,engine= 'xlrd')
on13 = pd.read_excel(on13_p,engine= 'xlrd')
frame = [on7, on13]
on_sales_master = pd.concat(frame,ignore_index=True)

sv7 = pd.read_excel(sv7_p,engine= 'xlrd')
sv13 = pd.read_excel(sv13_p,engine= 'xlrd')
frame = [sv7, sv13]
sv_sales_master = pd.concat(frame,ignore_index=True)

si7_inv = pd.read_excel(si7_inv_p,engine= 'xlrd')
si13_inv = pd.read_excel(si13_inv_p,engine= 'xlrd')
on7_inv = pd.read_excel(on7_inv_p,engine= 'xlrd')
on13_inv = pd.read_excel(on13_inv_p,engine= 'xlrd')
sv7_inv = pd.read_excel(sv7_inv_p,engine= 'xlrd')
sv13_inv = pd.read_excel(sv13_inv_p,engine= 'xlrd')

frame = [si7_inv, si13_inv]
si_inv_master = pd.concat(frame,ignore_index=True)

frame = [on7_inv, on13_inv]
on_inv_master = pd.concat(frame,ignore_index=True)

frame = [sv7_inv, sv13_inv]
sv_inv_master = pd.concat(frame,ignore_index=True)

inv_rep = pd.read_excel(iv_rep,header=2)

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
si_sales_master['amount'] = si_sales_master['amount'].astype(float)

################################################### Creating Master sheets #############################################
frame = [si_sales_master,on_sales_master,sv_sales_master]
sales_master = pd.concat(frame,ignore_index=True)

frame = [si_inv_master,on_inv_master,sv_inv_master]
inv_master = pd.concat(frame,ignore_index=True)


####################################################### Changes ########################################################
# DF merge to obtain CBM and CasePack from INV for SalesMaster(VLOOKUP SIM)
sales_master["ID"] = sales_master["style"]+sales_master["color"]

inv_rep = inv_rep[['ID','CRTN CBM','Casepack']]

sales_master = pd.merge(sales_master,inv_rep,on='ID',how='inner')

inv_master = inv_master[['WH','style','grp','desc','color','division','cubic_ft','weight','master_pack','unit',
                               'caseqty', 'cubic']]

# Rounding
sales_master['Total CBM'] = round((sales_master['CRTN CBM']/sales_master['Casepack'])*sales_master['shptot'],4)
sales_master['CRTN CBM'] = round(sales_master['CRTN CBM'],4)

#### Summary sheets
########################################### Sales by wh by month #######################################################

sales_month = pd.DataFrame()
list = ['01','02','03','04','05','06','07','08','09','10','11','12','']
sales_month_master = pd.DataFrame(list,columns=['Month'])

sales_month = sales_master[sales_master.WH =='ON']
sales_month = sales_month[['WH','Month','amount']]
sales_month_on= sales_month[sales_month.Month =='01']
on_01 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='02']
on_02 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='03']
on_03 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='04']
on_04 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='05']
on_05 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='06']
on_06 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='07']
on_07 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='08']
on_08 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='09']
on_09 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='10']
on_10 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='11']
on_11 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='12']
on_12 = sum(sales_month_on['amount'])

on_sum = on_01+on_02+on_03+on_04+on_05+on_06+on_07+on_08+on_09+on_10+on_11+on_12
list = [on_01,on_02,on_03,on_04,on_05,on_06,on_07,on_08,on_09,on_10,on_11,on_12,on_sum]
sales_month_master['ON'] = list

sales_month = sales_master[sales_master.WH =='SI']
sales_month = sales_month[['WH','Month','amount']]
sales_month_on= sales_month[sales_month.Month =='01']
si_01 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='02']
si_02 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='03']
si_03 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='04']
si_04 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='05']
si_05 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='06']
si_06 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='07']
si_07 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='08']
si_08 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='09']
si_09 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='10']
si_10 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='11']
si_11 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='12']
si_12 = sum(sales_month_on['amount'])

si_sum = si_01+si_02+si_03+si_04+si_05+si_06+si_07+si_08+si_09+si_10+si_11+si_12
list = [si_01,si_02,si_03,si_04,si_05,si_06,si_07,si_08,si_09,si_10,si_11,si_12,si_sum]
sales_month_master['SI'] = list

sales_month = sales_master[sales_master.WH =='SV']
sales_month = sales_month[['WH','Month','amount']]
sales_month_on= sales_month[sales_month.Month =='01']
sv_01 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='02']
sv_02 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='03']
sv_03 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='04']
sv_04 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='05']
sv_05 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='06']
sv_06 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='07']
sv_07 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='08']
sv_08 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='09']
sv_09 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='10']
sv_10 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='11']
sv_11 = sum(sales_month_on['amount'])
sales_month_on= sales_month[sales_month.Month =='12']
sv_12 = sum(sales_month_on['amount'])

sv_sum = sv_01+sv_02+sv_03+sv_04+sv_05+sv_06+sv_07+sv_08+sv_09+sv_10+sv_11+sv_12
list = [sv_01,sv_02,sv_03,sv_04,sv_05,sv_06,sv_07,sv_08,sv_09,sv_10,sv_11,sv_12,sv_sum]
sales_month_master['SV'] = list

tot01 = on_01 + si_01 + sv_01
tot02 = on_02 + si_02 + sv_02
tot03 = on_03 + si_03 + sv_03
tot04 = on_04 + si_04 + sv_04
tot05 = on_05 + si_05 + sv_05
tot06 = on_06 + si_06 + sv_06
tot07 = on_07 + si_07 + sv_07
tot08 = on_08 + si_08 + sv_08
tot09 = on_09 + si_09 + sv_09
tot10 = on_10 + si_10 + sv_10
tot11 = on_11 + si_11 + sv_11
tot12 = on_12 + si_12 + sv_12
totsum = on_sum + si_sum + sv_sum
list = [tot01,tot02,tot03,tot04,tot05,tot06,tot07,tot08,tot09,tot10,tot11,tot12,totsum]
sales_month_master['Grand Total'] = list


#################################### Units by WH by month ##############################################################
unit_month = pd.DataFrame()
list = ['01','02','03','04','05','06','07','08','09','10','11','12','']
unit_month_master = pd.DataFrame(list,columns=['Month'])

unit_month = sales_master[sales_master.WH =='ON']
unit_month = unit_month[['WH','Month','shptot']]
unit_month_on= unit_month[unit_month.Month =='01']
on_01 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='02']
on_02 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='03']
on_03 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='04']
on_04 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='05']
on_05 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='06']
on_06 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='07']
on_07 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='08']
on_08 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='09']
on_09 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='10']
on_10 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='11']
on_11 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='12']
on_12 = sum(unit_month_on['shptot'])

on_sum = on_01+on_02+on_03+on_04+on_05+on_06+on_07+on_08+on_09+on_10+on_11+on_12
list = [on_01,on_02,on_03,on_04,on_05,on_06,on_07,on_08,on_09,on_10,on_11,on_12,on_sum]
unit_month_master['ON'] = list

unit_month = sales_master[sales_master.WH =='SI']
unit_month = unit_month[['WH','Month','shptot']]
unit_month_on= unit_month[unit_month.Month =='01']
si_01 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='02']
si_02 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='03']
si_03 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='04']
si_04 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='05']
si_05 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='06']
si_06 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='07']
si_07 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='08']
si_08 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='09']
si_09 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='10']
si_10 = sum(unit_month_on['shptot'])
sales_month_on= unit_month[unit_month.Month =='11']
si_11 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='12']
si_12 = sum(unit_month_on['shptot'])

si_sum = si_01+si_02+si_03+si_04+si_05+si_06+si_07+si_08+si_09+si_10+si_11+si_12
list = [si_01,si_02,si_03,si_04,si_05,si_06,si_07,si_08,si_09,si_10,si_11,si_12,si_sum]
unit_month_master['SI'] = list

unit_month = sales_master[sales_master.WH =='SV']
unit_month = unit_month[['WH','Month','shptot']]
unit_month_on= unit_month[unit_month.Month =='01']
sv_01 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='02']
sv_02 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='03']
sv_03 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='04']
sv_04 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='05']
sv_05 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='06']
sv_06 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='07']
sv_07 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='08']
sv_08 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='09']
sv_09 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='10']
sv_10 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='11']
sv_11 = sum(unit_month_on['shptot'])
unit_month_on= unit_month[unit_month.Month =='12']
sv_12 = sum(unit_month_on['shptot'])

sv_sum = sv_01+sv_02+sv_03+sv_04+sv_05+sv_06+sv_07+sv_08+sv_09+sv_10+sv_11+sv_12
list = [sv_01,sv_02,sv_03,sv_04,sv_05,sv_06,sv_07,sv_08,sv_09,sv_10,sv_11,sv_12,sv_sum]
unit_month_master['SV'] = list

tot01 = on_01 + si_01 + sv_01
tot02 = on_02 + si_02 + sv_02
tot03 = on_03 + si_03 + sv_03
tot04 = on_04 + si_04 + sv_04
tot05 = on_05 + si_05 + sv_05
tot06 = on_06 + si_06 + sv_06
tot07 = on_07 + si_07 + sv_07
tot08 = on_08 + si_08 + sv_08
tot09 = on_09 + si_09 + sv_09
tot10 = on_10 + si_10 + sv_10
tot11 = on_11 + si_11 + sv_11
tot12 = on_12 + si_12 + sv_12
totsum = on_sum + si_sum + sv_sum
list = [tot01,tot02,tot03,tot04,tot05,tot06,tot07,tot08,tot09,tot10,tot11,tot12,totsum]
unit_month_master['Grand Total'] = list

############################################### Creating CBM by WH by mth ##############################################
cbm_month = pd.DataFrame()
list = ['01','02','03','04','05','06','07','08','09','10','11','12','']
cbm_month_master = pd.DataFrame(list,columns=['Month'])

cbm_month = sales_master[sales_master.WH =='ON']
cbm_month = cbm_month[['WH','Month','Total CBM']]
cbm_month_on= cbm_month[cbm_month.Month =='01']
on_01 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='02']
on_02 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='03']
on_03 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='04']
on_04 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='05']
on_05 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='06']
on_06 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='07']
on_07 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='08']
on_08 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='09']
on_09 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='10']
on_10 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='11']
on_11 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='12']
on_12 = sum(cbm_month_on['Total CBM'])

on_sum = on_01+on_02+on_03+on_04+on_05+on_06+on_07+on_08+on_09+on_10+on_11+on_12
list = [on_01,on_02,on_03,on_04,on_05,on_06,on_07,on_08,on_09,on_10,on_11,on_12,on_sum]
cbm_month_master['ON'] = list

cbm_month = sales_master[sales_master.WH =='SI']
cbm_month = cbm_month[['WH','Month','Total CBM']]
cbm_month_on= cbm_month[cbm_month.Month =='01']
si_01 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='02']
si_02 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='03']
si_03 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='04']
si_04 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='05']
si_05 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='06']
si_06 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='07']
si_07 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='08']
si_08 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='09']
si_09 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='10']
si_10 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='11']
si_11 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='12']
si_12 = sum(cbm_month_on['Total CBM'])

si_sum = si_01+si_02+si_03+si_04+si_05+si_06+si_07+si_08+si_09+si_10+si_11+si_12
list = [si_01,si_02,si_03,si_04,si_05,si_06,si_07,si_08,si_09,si_10,si_11,si_12,si_sum]
cbm_month_master['SI'] = list

cbm_month = sales_master[sales_master.WH =='SV']
cbm_month =cbm_month[['WH','Month','Total CBM']]
cbm_month_on= cbm_month[cbm_month.Month =='01']
sv_01 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='02']
sv_02 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='03']
sv_03 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='04']
sv_04 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='05']
sv_05 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='06']
sv_06 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='07']
sv_07 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='08']
sv_08 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='09']
sv_09 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='10']
sv_10 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='11']
sv_11 = sum(cbm_month_on['Total CBM'])
cbm_month_on= cbm_month[cbm_month.Month =='12']
sv_12 = sum(cbm_month_on['Total CBM'])

sv_sum = sv_01+sv_02+sv_03+sv_04+sv_05+sv_06+sv_07+sv_08+sv_09+sv_10+sv_11+sv_12
list = [sv_01,sv_02,sv_03,sv_04,sv_05,sv_06,sv_07,sv_08,sv_09,sv_10,sv_11,sv_12,sv_sum]
cbm_month_master['SV'] = list

tot01 = on_01 + si_01 + sv_01
tot02 = on_02 + si_02 + sv_02
tot03 = on_03 + si_03 + sv_03
tot04 = on_04 + si_04 + sv_04
tot05 = on_05 + si_05 + sv_05
tot06 = on_06 + si_06 + sv_06
tot07 = on_07 + si_07 + sv_07
tot08 = on_08 + si_08 + sv_08
tot09 = on_09 + si_09 + sv_09
tot10 = on_10 + si_10 + sv_10
tot11 = on_11 + si_11 + sv_11
tot12 = on_12 + si_12 + sv_12
totsum = on_sum + si_sum + sv_sum
list = [tot01,tot02,tot03,tot04,tot05,tot06,tot07,tot08,tot09,tot10,tot11,tot12,totsum]
cbm_month_master['Grand Total'] = list

############################################### Creating CBF by WH by mth ##############################################
cbf_month_master = pd.DataFrame()
#cbf_month_master = cbm_month_master.astype(float)
cbf_month_master['Month'] = cbm_month_master['Month']
cbf_month_master['ON'] = cbm_month_master['ON'] * 35.3146667
cbf_month_master['SI'] = cbm_month_master['SI'] * 35.3146667
cbf_month_master['SV'] = cbm_month_master['SV'] * 35.3146667
cbf_month_master['Grand Total'] = cbm_month_master['Grand Total'] * 35.3146667

############################################### Creating Unit by WH Total ##############################################

unit_tot_on = sum(on_inv_master['unit'])
unit_tot_si = sum(si_inv_master['unit'])
unit_tot_sv = sum(sv_inv_master['unit'])
unit_tot_tot = unit_tot_on+unit_tot_si+unit_tot_sv
unit_tot = pd.DataFrame({'': ["Sum of Unit"],'ON': [unit_tot_on],'SI': [unit_tot_si],'SV': [unit_tot_sv],'GrandTotal': [unit_tot_tot]})

############################################## Creating Cubic by WH Total ##############################################

cube_tot_on = sum(on_inv_master['cubic'])
cube_tot_si = sum(si_inv_master['cubic'])
cube_tot_sv = sum(sv_inv_master['cubic'])
cube_tot_tot = cube_tot_on+cube_tot_si+cube_tot_sv
cube_tot = pd.DataFrame({'': ["Sum of Cubic"],'ON': [cube_tot_on],'SI': [cube_tot_si],'SV': [cube_tot_sv],'GrandTotal': [cube_tot_tot]})
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

sales_month_master.to_excel(fileName, sheet_name='Sales by WH by Month', index = False)
unit_month_master.to_excel(fileName, sheet_name='Units by WH by Month', index = False)
cbm_month_master.to_excel(fileName, sheet_name='CBM by WH by Month', index = False)
cbf_month_master.to_excel(fileName, sheet_name='CBF by WH by Month', index = False)

unit_tot.to_excel(fileName, sheet_name='unit totals', index = False)
cube_tot.to_excel(fileName, sheet_name='cubic totals', index = False)

fileName.save()
