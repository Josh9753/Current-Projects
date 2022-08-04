import glob
import pandas as pd
import numpy as np
import datetime
import xlwt
import openpyxl
import math
from openpyxl import load_workbook

###################################################### DATA SOURCES ######################################################

styles_categories = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\1 - STYLES\Master.xlsx"
categories_reference = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\1 - STYLES\Categories_Master.xlsx"
sty_files = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\1 - STYLES\* - Style Report.xls"
bookings = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\13 - BOOKINGS\Bookings Approvals_Master.xlsx"
opos_files = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\3 - PURCHASE ORDERS\* - PO OR Comp Report.csv"
fopo_orders = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\3 - PURCHASE ORDERS\* - FPO.xls"
shi_files = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\4 - SHIPMENTS\* - PO Shipment Report.xls"
ord_files = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\7 - SALES ORDERS\* - Open Order.csv"
vh_files = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\01 - AIMS RAW DATA\6 - FACTORY ORDERS\* - Vendor Purchases Report.csv"
shipto = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\25 - PO TRACKING\01 - WORKING\ShipTo.csv"
gc = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\25 - PO TRACKING\Detailed-Tracking.xlsx" 
mxn = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\25 - PO TRACKING\Mexico Follow Up.xlsx"
wt = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\25 - PO TRACKING\WebTracker - *.xls"
cbf = r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\05 - INVENTORY\01 - HISTORICAL\CBF.xlsx"

######################################################## GET DATA ########################################################

#styles categories master
styles_categories_data = pd.read_excel(styles_categories, na_values="NA")
category = pd.read_excel(categories_reference, na_values="NA")
category = category.rename(index = str, columns = {"category":"Category", "division":"Division"})

#styles master
st_master_files = glob.glob(sty_files)
st_master_reports = []
for file in st_master_files: 
    st_master_data = pd.read_excel(file, na_values="NA",engine = 'xlrd')
    st_master_data["data"] = file.split(" ")[-4][7:]
    st_master_reports.append(st_master_data)
st_master = pd.concat(st_master_reports, sort=False)

st_master["Database_Number"] = st_master["data"].astype("int")
st_master["Style"] = st_master["style"].astype("str")
st_master["Color"] = st_master["color"].astype("str")
st_master["Style_Description"] = st_master["stydesc"].astype("str")
st_master["Color_Description"] = st_master["clrdesc"].astype("str")
st_master["AIMS_Division"] = st_master["division"].astype("str")
st_master["PO_Cost"] = st_master["pocost"].astype("float")
st_master["In_Land_Cost"] = st_master["inlandftr"].astype("float")
st_master["Ocean_Freight_Cost"] = st_master["oceanfrt"].astype("float")
st_master["Duty"] = st_master["duty"].astype("float")
st_master["Tariff"] = st_master["tariff"].astype("float")
st_master["Other_Cost"] = st_master["othercost"].astype("float")
st_master["Total_Cost"] = st_master["totalcost"].astype("float")
st_master["Sales_Price"] = st_master["salesprice"].astype("float")
st_master["Unit_Weight"] = st_master["unitweight"].astype("float")
st_master["Lenght"] = st_master["length"].astype("float")
st_master["Widht"] = st_master["width"].astype("float")
st_master["Height"] = st_master["height"].astype("float")
st_master["Caseweight"] = st_master["caseweight"].astype("float")
st_master["Casepack"] = st_master["casepack"].astype("int")
st_master["UPC"] = st_master["upcno"].astype("str")
st_master["Date_Created"] = st_master["crdate"].astype("str")

styles_data = st_master[["Database_Number","Style","Color","Style_Description","Color_Description","AIMS_Division",
                "PO_Cost","In_Land_Cost","Ocean_Freight_Cost","Duty","Tariff","Other_Cost","Total_Cost","Sales_Price",
                "Unit_Weight","Lenght","Widht","Height","Caseweight","Casepack","UPC","Date_Created"]]

#pos
pos_files = glob.glob(opos_files)
pos_reports = []
for file in pos_files: 
    pos_data = pd.read_csv(file, na_values=0, encoding = "latin-1")
    pos_data["data"] = file.split(" ")[-6][7:]
    pos_reports.append(pos_data)
opos_master = pd.concat(pos_reports)

opos_master["Database_Number"] = opos_master["data"].astype("int")
opos_master["Vendor"] = opos_master["vendor"].astype("str")
opos_master = opos_master[opos_master["Vendor"] != "nan"]
opos_master["Vendor_Name"] = opos_master["vname"].astype("str")
opos_master["PO_No"] = opos_master["pono"].astype("int")
opos_master["Style"] = opos_master["style"].astype("str")
opos_master["Color"] = opos_master["color"].astype("str")
opos_master["PO_Date"] = pd.to_datetime(opos_master["podate"])
opos_master["PO_Cancel"] = pd.to_datetime(opos_master["pocancel"])
opos_master["PO_Qty"] = opos_master["poqty"].replace(np.nan,0).astype("int")
opos_master["Shipped_Qty"] = opos_master["shpqty"].replace(np.nan,0).astype("int")
opos_master["Open_Qty"] = opos_master["openqty"].replace(np.nan,0).astype("int")
opos_master["FOB"] = opos_master["fob"].astype("str")
opos_master["Shipment_No"] = opos_master["shipno"].fillna(0).astype("int")

pos_data = opos_master[["Database_Number","Style","Color","Vendor_Name","PO_No", "PO_Qty", "PO_Date", "PO_Cancel","Shipped_Qty","Open_Qty","FOB","Shipment_No"]] 

#shipments
ship_files = glob.glob(shi_files)
ship_reports = []
for file in ship_files: 
    ship_data = pd.read_excel(file, na_values=0,engine = 'xlrd')
    ship_data["data"] = file.split(" ")[10][10:] #[9][10:]
    ship_reports.append(ship_data)
shi_master = pd.concat(ship_reports, sort=False)


shi_master["Database_Number"] = shi_master["data"].astype("int")
shi_master["Style"] = shi_master["style"].astype("str")
shi_master["Color"] = shi_master["color"].astype("str")
shi_master["Shipment_No"] = shi_master["shipment_no"].astype("str")
shi_master["Document_No"] = shi_master["document_num"].astype("str")
shi_master["Container"] = shi_master["container"].astype("str")
shi_master["PO_No"] = shi_master["pono"].astype("int")
shi_master["ETD_Port"] = pd.to_datetime((shi_master["shipment_date"]).astype("str").str.replace("  -   -","01-Jan-1900"))
shi_master["ETA_Port"] = pd.to_datetime(shi_master["complete_date"].astype("str").str.replace("  -   -","01-Jan-1900"))
shi_master["PO_Cost"] = shi_master["purchase_cost"].astype("float").fillna(0)
shi_master["Duty"] = shi_master["duty_cost"].astype("float").fillna(0)
shi_master["Ocean_Freight"] = shi_master["ocean_frght"].astype("float").fillna(0)
shi_master["In_Land_Freight"] = shi_master["inland_frght"].astype("float").fillna(0)
shi_master["Tariff"] = shi_master["tariff"].astype("float").fillna(0)
shi_master["Misc_1"] = shi_master["misc_1nd"].astype("float").fillna(0)
shi_master["Misc_2"] = shi_master["misc_2nd"].astype("float").fillna(0)
shi_master["Misc_4"] = shi_master["misc_4d"].astype("float").fillna(0)
shi_master["Misc_5"] = shi_master["misc_5d"].astype("float").fillna(0)
shi_master["Style_Cost"] = shi_master["style_cost"].astype("float").fillna(0)
shi_master["Issued_Qty"] = shi_master["issue_qty"].astype("int").fillna(0)
shi_master["Recieved_Qty"] = shi_master["recvd_qty"].astype("int").fillna(0) 
shi_master["Recieved_Cost"] = shi_master["recvd_cost"].astype("float").fillna(0)

shipping_data = shi_master[["Database_Number","Style","Color","Shipment_No","Document_No","Container","PO_No","ETD_Port","ETA_Port",
                "PO_Cost","Duty","Ocean_Freight","In_Land_Freight","Tariff","Misc_1","Misc_2","Misc_4","Misc_5",
                "Style_Cost","Issued_Qty","Recieved_Qty","Recieved_Cost"]] 

#vendor history
fty_files = glob.glob(vh_files)
fty_reports = []
for file in fty_files: 
    fty_data = pd.read_csv(file, na_values=0, encoding = "latin-1")
    fty_data["data"] = file.split(" ")[-5][7:]
    fty_reports.append(fty_data)
fty_master = pd.concat(fty_reports, sort=False)

fty_master["Database_Number"] = fty_master["data"].astype("int")
fty_master["Style"] = fty_master["style"].astype("str")
fty_master["Color"] = fty_master["color"].astype("str")
fty_master["eta"] = fty_master["eta"].str.replace("/  /","NA")
#fty_master = fty_master[fty_master["eta"] != "NA"]
fty_master = fty_master[fty_master["complete"] != '01/25/8201']
fty_master["Group_Code"] = fty_master["group"].astype("str")
fty_master = fty_master[fty_master["Group_Code"] != "LO"]
fty_master = fty_master[fty_master["Group_Code"] != "ST"]
fty_master = fty_master[fty_master["Group_Code"] != "RD"]
fty_master["Vendor"] = fty_master["vendor"].astype("str")
fty_master["Vendor_Name"] = fty_master["vendname"].astype("str")
fty_master["Account"] = fty_master["account"].astype("str")
fty_master["Account_Name"] = fty_master["acctname"].astype("str")
fty_master["PO_No"] = fty_master["pono"].astype("int")
fty_master["Issued"] = pd.to_datetime(fty_master["issued"])
fty_master["Year_Issued"] = fty_master["Issued"].dt.year
fty_master = fty_master[fty_master["Year_Issued"] > 2015]
fty_master["Complete"] = pd.to_datetime(fty_master["complete"])
fty_master["Ordered"] = fty_master["ordered"].replace(np.nan,0).astype("int")
fty_master["Amount"] = fty_master["amt"].replace(np.nan,0).astype("float")
fty_master["Received"] = fty_master["received"].replace(np.nan,0).astype("int")
fty_master["Damaged"] = fty_master["damage"].replace(np.nan,0).astype("int")
fty_master["Balance"] = fty_master["balance"].replace(np.nan,0).astype("int")

vendor_history_data = fty_master[["Database_Number","Style","Color","Group_Code","Vendor","Vendor_Name","Account","Account_Name","PO_No",
                 "Issued","Complete","Ordered","Amount","Received","Damaged","Balance"]]

#factory open orders
fopo_files = glob.glob(fopo_orders)
fopo_reports = []
for file in fopo_files: 
    fopo_data = pd.read_excel(file, na_values=0,engine = 'xlrd')
    fopo_data["data"] = file.split(" ")[-3][7:]
    fopo_reports.append(fopo_data)
fopo_report = pd.concat(fopo_reports)
fopo_report = fopo_report[["pono","loc","custpo"]]
fopo_report = fopo_report.rename(index = str, columns = {"pono":"PO No","loc":"AIMS Location Code"})
fopo_report = fopo_report.drop_duplicates(subset="PO No", keep="first")

#Ship to
shipto_data = pd.read_csv(shipto, na_values=0, encoding = "latin-1")

#bookings
bkngs = pd.read_excel(bookings, sheet_name="Approved Bookings")
bkngs = bkngs[["UNIQUE ID","Booking #","PO #","Approval Date","Remarks"]]
bkngs = bkngs.rename(index = str, columns = {"UNIQUE ID":"id","Booking #":"Booking No","PO #":"PO No",
                                            "Approval Date":"Shipping Approval Date","Remarks":"PT Remarks"})
bkngs["PT Remarks"] = bkngs["PT Remarks"].fillna(" ")
bkngs["Shipping Approval Date"] = pd.to_datetime(bkngs["Shipping Approval Date"])
bkngs = bkngs.sort_values(by=["id","Shipping Approval Date"], ascending=[True,True])
bkngs = bkngs.drop_duplicates(subset=["id","PO No"], keep="last")
bkngs["poid"] = bkngs["id"].astype("str") + bkngs["PO No"].astype("str")
bkngs["bkngid"] = bkngs["PO No"].astype("str")# + bkngs["Booking No"].astype("str") #id by PO only
#bkngs["bkngid"] = bkngs["PO No"].astype("str") + bkngs["Booking No"].astype("str") 
bkngs = bkngs[["poid","bkngid","Booking No"]]

#GoComet
gc_data = pd.read_excel(gc, na_values="NA")
gc_data = gc_data[["Tracking Number","Carrier Name","Origin Departure Vessel Name","Origin Port","Destination Port",
                   "Status","Origin Departure Planned Date (ETD)","Origin Departure Actual Date",
                   "Destination Arrival Planned Date (ETA)","Destination Arrival Actual Date"]] 
gc_data = gc_data.rename(index = str, columns = {"Tracking Number":"Tracking Container No",
                                                "Status":"Shipment Status"})

#Narin's File
mxntr = pd.read_excel(mxn, na_values="NA",engine = 'openpyxl')
mxntr = mxntr[["PO#","Doc#"," BK#","CNTR#","ETD","ETA LZC / Discharge","Est WH DATE"]]
mxntr["Doc#"] = mxntr["Doc#"].fillna("NA")
mxntr = mxntr[mxntr["Doc#"] != "NA"]
mxntr["CNTR#"] = mxntr["CNTR#"].str.replace("\t\t\t","")
mxntr = mxntr.fillna(method='ffill')
mxntr["PO#"] = mxntr["PO#"].astype("str")
mxntr["PO#"] = mxntr["PO#"].apply(lambda x: x.split(' '))
mxntr = mxntr.explode("PO#")
mxntr["PO#"] = mxntr["PO#"].str.replace(",","")
mxntr[" BK#"] = mxntr[" BK#"].fillna("NA")
mxntr = mxntr[mxntr[" BK#"] != "NA"]
mxntr = mxntr[mxntr["CNTR#"] != "BOOKING UPDATE"]
mxntr["bkngid"] = mxntr["PO#"].fillna(0).astype("str")# + mxntr[" SO#"].fillna(0).astype("str") #id by PO only
#mxntr["bkngid"] = mxntr["PO#"].fillna(0).astype("str") + mxntr[" SO#"].fillna(0).astype("str")
mxntr = mxntr.drop_duplicates(subset=["PO#","Doc#"," BK#","CNTR#"], keep="first")
mxntr = mxntr[["bkngid","CNTR#","ETD","ETA LZC / Discharge","Est WH DATE"]]
mxntr = mxntr.rename(index = str, columns = {"CNTR#":"Import Container No","ETD":"Import ETD Port",
                                             "ETA LZC / Discharge":"Import ETA Port","Est WH DATE":"Import ETA Whs"})

#Web Tracker
files = glob.glob(wt)
reports = []
for file in files: 
    data = pd.read_excel(file, na_values=0,engine = 'xlrd')
    reports.append(data)
wt_report = pd.concat(reports, sort=False)

wt_report = wt_report[["Shipment#","ETD","ETA","Destination","Containers","Order Ref#","Shipper's Ref#"]]

wt_report["ETD"] = pd.to_datetime(wt_report["ETD"]).dt.date
wt_report["Year"] = (pd.to_datetime(wt_report["ETD"]).dt.year).fillna(0).astype("int")
wt_report = wt_report[wt_report["Year"] > 2020]
wt_report["ETA"] = pd.to_datetime(wt_report["ETA"]).dt.date
wt_report["Containers"] = wt_report["Containers"].str.replace("\r\n","")

wtcl = wt_report
wtcl["Order Ref#"] = wtcl["Order Ref#"].astype("str")
wtcl["Order Ref#"] = wtcl["Order Ref#"].apply(lambda x: x.split(' '))
wtcl = wtcl.explode("Order Ref#")
wtcl["Order Ref#"] = wtcl["Order Ref#"].str.replace(",","")
wtcl = wtcl.rename(index=str, columns={"Order Ref#":"PO No","Containers":"WT Container(s)","Shipment#":"Booking No",
                                      "ETD":"WT Port ETD","ETA":"WT Port ETA"})
wtcl["bkngid"] = wtcl["PO No"].fillna(0).astype("str")# + wtcl["Booking No"].fillna(0).astype("str") #id by PO only
#wtcl["bkngid"] = wtcl["PO No"].fillna(0).astype("str") + wtcl["Booking No"].fillna(0).astype("str")
wt_data = wtcl[["bkngid","WT Container(s)","WT Port ETD","WT Port ETA"]]

#CBF for sets
cbf_data = pd.read_excel(cbf, na_values="NA")

####################################################### CLEAN DATA #######################################################

#STYLES
#filtering replenishment styles by status & type 
styles_categories_data["ID"] = styles_categories_data["STYLE"].astype("str") + styles_categories_data["COLOR"].astype("str")
categories = styles_categories_data[["DATABASE_NUMBER","DIVISION","ID","PRODUCT_CATEGORY","PRODUCT_TYPE","PRODUCT_STATUS"]]
categories = categories.rename(index = str, columns = {"DATABASE_NUMBER":"AIMS Database", "DIVISION":"AIMS Division",
                                                       "PRODUCT_CATEGORY":"Category", "PRODUCT_TYPE":"Type",
                                                       "PRODUCT_STATUS":"Status"})
categories = categories[categories["Category"] != "Curtain Rod"]
categories = categories[categories["Category"] != "Locks"]
categories = categories[categories["Category"] != "Storage"]
categories = categories[categories["Category"] != "Misc"]
categories = categories[categories["Category"] != "Others - Enchante Acc"]
categories = categories[categories["AIMS Database"] != 10] # March 26, 2021
categories = pd.merge(categories, category, left_on = "Category", right_on = "Category", how = "left")

styles_filter = categories

#styles details
styles_data["ID"] = styles_data["Style"].astype("str") + styles_data["Color"].astype("str")
styles_data = styles_data[styles_data["Database_Number"] != 10]
styles_data["Weight KG"] = styles_data["Caseweight"] * 0.45359237
styles_data["Lenght CM"] = styles_data["Lenght"] * 2.54
styles_data["Width CM"] = styles_data["Widht"] * 2.54
styles_data["Height CM"] = styles_data["Height"] * 2.54
styles_data["CRTN CBM"] = styles_data["Lenght CM"] * styles_data["Width CM"] * styles_data["Height CM"] / 1000000
styles_data = styles_data[["ID", "Style", "Color", "Total_Cost","Sales_Price", "Lenght CM", "Width CM", "Height CM", "CRTN CBM", 
                           "Weight KG", "Casepack", "Date_Created"]]
styles_data = styles_data.drop_duplicates("ID", keep="first")
styles_data = styles_data.rename(index = str, columns = {"Total_Cost":"ELC", "Sales_Price":"Sales Price",
                                                         "Date_Created":"Date Created"})
styles_master = pd.merge(categories, styles_data, left_on = "ID", right_on = "ID", how = "left")
styles_master = styles_master[["Division","Category","ID","AIMS Database","AIMS Division","Style","Color","Type","Status",
                              "ELC","Sales Price"]]

dimens = styles_data[["ID","Lenght CM","Width CM","Height CM","CRTN CBM","Casepack"]]
cbf_data["CBM"] = cbf_data["CBF"]/35.214667
cbf_data["Audit"] = "Y"
dimens = pd.merge(dimens, cbf_data, left_on = "ID", right_on = "ID", how = "left")
dimens.loc[:,"CRTN CBM"][dimens.loc[:,"Audit"] == "Y"] = dimens.loc[:,"CBM"]
dimens.loc[:,"Casepack"][dimens.loc[:,"Audit"] == "Y"] = dimens.loc[:,"CP"]
del dimens["CP"]
del dimens["CBF"]
del dimens["CBM"]
del dimens["Audit"]
dimens = dimens.drop_duplicates("ID", keep="first")

#PURCHASE ORDERS
pos_data["ID"] = pos_data["Style"] + pos_data["Color"]
pos_data["POID"] = pos_data["ID"] + pos_data["PO_No"].astype("str") + pos_data["Shipment_No"].astype("str")
pos_data = pos_data[["POID", "ID", "Style", "Color", "PO_No", "PO_Date", "PO_Cancel", "PO_Qty", "Vendor_Name","FOB"]]

#SHIPMENT
shipping_data["ID"] = shipping_data["Style"] + shipping_data["Color"]
shipping_data["POID"] = shipping_data["ID"] + shipping_data["PO_No"].astype("str") + shipping_data["Shipment_No"].astype("str")
shipping_data["In_Transit"] = shipping_data["Issued_Qty"] - shipping_data["Recieved_Qty"]
shipping_data = shipping_data[["POID", "Shipment_No", "Document_No", "Container", "ETD_Port", 
                               "ETA_Port", "Issued_Qty", "Recieved_Qty", "In_Transit"]]

#POs DATA
pos = pd.merge(pos_data, shipping_data, left_on = "POID", right_on = "POID", how = "left")
pos["Container"] = pos["Container"].fillna("NA")
pos = pos.drop_duplicates(subset=["POID", "Issued_Qty", "Container"], keep='first')
pos["Issued_Qty"] = pos["Issued_Qty"].fillna(0).astype("int")
pos["Recieved_Qty"] = pos["Recieved_Qty"].fillna(0).astype("int")
pos["In_Transit"] = pos["In_Transit"].fillna(0).astype("int")
pos["PO_Date"] = pd.to_datetime(pos["PO_Date"])
pos["PO_Cancel"] = pd.to_datetime(pos["PO_Cancel"])
pos["ETD_Port"] = pd.to_datetime(pos["ETD_Port"])
pos["ETA_Port"] = pd.to_datetime(pos["ETA_Port"])
pos = pos[pos["Recieved_Qty"] == 0]

#coming inventory
coming_inv = pos
coming_inv.loc[:,"Est. ETA Whs"] = coming_inv.loc[:,"PO_Cancel"] + pd.DateOffset(days = 45)
coming_inv.loc[:,"ETA Whs"] = coming_inv.loc[:,"ETA_Port"] + pd.DateOffset(days = 10)
coming_inv.loc[:,"ETA Whs"] = coming_inv.loc[:,"ETA Whs"].fillna(0)
coming_inv.loc[coming_inv["ETA Whs"]==0,"ETA Whs"] = coming_inv["Est. ETA Whs"]
coming_inv.loc[:,"Qty"] = coming_inv.loc[:,"In_Transit"]
coming_inv.loc[coming_inv["Qty"]==0,"Qty"] = coming_inv["PO_Qty"]
coming_inv["ETD_Port"] = coming_inv["ETD_Port"].fillna("NA")
coming_inv["ETA_Port"] = coming_inv["ETA_Port"].fillna("NA")
coming_inv["Shipment_No"] = coming_inv["Shipment_No"].fillna("NA")
coming_inv["Document_No"] = coming_inv["Document_No"].fillna("NA")
coming_inv["ETA Whs"] = pd.to_datetime(coming_inv["ETA Whs"])
coming_inv.loc[:,"Week No"] = coming_inv["ETA Whs"].dt.week.astype("str")
coming_inv.loc[:,"Year"] = coming_inv.loc[:,"ETA Whs"].dt.year.astype("str")
coming_inv.loc[:,"Week Date"] = (coming_inv.loc[:,"Year"] + "-W" + coming_inv.loc[:,"Week No"]).astype("str")
coming_inv.loc[:,"Week"] = [datetime.datetime.strptime(x + '-1', "%Y-W%W-%w") for x in coming_inv.loc[:,"Week Date"]]
coming_inventory = coming_inv[["ID", "POID", "Style", "Color", "PO_No", "PO_Date", "PO_Cancel", "Vendor_Name","FOB", 
                               "Shipment_No", "Document_No", "Container", "Qty", "ETD_Port", "ETA_Port", "ETA Whs", "Week"]]
coming_inventory.loc[:,"Week"] = pd.to_datetime(coming_inventory.loc[:,"Week"])
today = datetime.datetime.today() - pd.DateOffset(days = 7)

#coming POs
list_of_POs = coming_inventory[coming_inventory["Qty"] != 0]
list_of_POs = pd.merge(styles_filter, list_of_POs, left_on = "ID", right_on = "ID", how = "left")
list_of_POs["Qty"] = list_of_POs["Qty"].fillna(0)
list_of_POs = list_of_POs[list_of_POs["Qty"] != 0]
del list_of_POs["POID"]

#master
list_of_POs = pd.merge(list_of_POs, dimens, left_on = "ID", right_on = "ID", how = "left")
list_of_POs["CRTNs"] = list_of_POs["Qty"] / list_of_POs["Casepack"]
list_of_POs["Total CBM"] = list_of_POs["CRTNs"] * list_of_POs["CRTN CBM"]
list_of_POs["Total CBF"] = list_of_POs["Total CBM"] * 35.214667
list_of_POs["CRTNs"] = list_of_POs["Qty"] / list_of_POs["Casepack"]

bound_in = list_of_POs[["Division","Category","ID","AIMS Database","AIMS Division","Style","Color","Type","Status",
                       "PO_No","PO_Date","PO_Cancel","Vendor_Name","Shipment_No","Document_No","Container",
                       "Qty","Casepack","CRTNs","CRTN CBM","Total CBM","Total CBF","ETD_Port","ETA_Port","FOB"]]
#bound_in = bound_in[bound_in["Division"] != "Bathroom Acc"]
bound_in = bound_in.sort_values(by=["Division","Category","AIMS Database","ID"], ascending=[True,True,True,True])
bound_in["Audit"] = bound_in["ID"] + bound_in["PO_No"].astype("str") + bound_in["Qty"].astype("str") + bound_in["Container"].astype("str")
bound_in = bound_in.drop_duplicates(["Audit"], keep="first")
del bound_in["Audit"]
bound_in = bound_in.rename(index = str, columns = {"PO_No":"PO No","PO_Date":"PO Date","PO_Cancel":"PO Cancel",
                                                   "Vendor_Name":"Vendor Name","Shipment_No":"Shipment No",
                                                   "Document_No":"Document No","ETD_Port":"AIMS ETD Port",
                                                   "ETA_Port":"AIMS ETA Port","Container":"Container No"})
#port cleaning
bound_in.loc[:,"Vendor Port"] = " "
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Ningbo"] = "Ningbo, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "NINBO"] = "Ningbo, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "NINGBO"] = "Ningbo, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Sehnzhen"] = "Shenzhen, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Shenzhen"] = "Shenzhen, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "SHENZHEN"] = "Shenzhen, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Shanghai"] = "Shanghai, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "SHANGHAI"] = "Shanghai, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Tianjin"] = "Tianjin, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "TIANJIN"] = "Tianjin, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Xiamen"] = "Xiamen, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "XIAMEN"] = "Xiamen, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "Yantian"] = "Yantian, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "YANTIAN"] = "Yantian, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "YANTINA"] = "Yantian, CN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "KILKATA"] = "Kolkata, IN"
bound_in.loc[:,"Vendor Port"][bound_in.loc[:,"FOB"] == "KOLKATA"] = "Kolkata, IN"

master = pd.merge(bound_in, fopo_report, left_on = "PO No", right_on = "PO No", how = "left")
master = pd.merge(master, shipto_data, left_on = "AIMS Location Code", right_on = "AIMS Location Code", how = "left")
master.loc[:,"Delivery Port / Country"][(master.loc[:,"Vendor Port"] == "Kolkata, IN") & (master.loc[:,"Ship to Location"] == "FOB")] = "FOB Origin, IN"

master = master[["Division","Category","ID","AIMS Database","AIMS Division","Style","Color","Type","Status",
                 "PO No","PO Date","PO Cancel","Vendor Name","Shipment No","Document No","Container No","Qty",
                 "Casepack","CRTNs","CRTN CBM","Total CBM","Total CBF","AIMS ETD Port","AIMS ETA Port","Vendor Port",
                 "AIMS Location Code","Ship to Location","Delivery Port / Country"]]
master["poid"] = master["ID"].astype("str") + master["PO No"].astype("int").astype("str")
master = pd.merge(master, bkngs, left_on = "poid", right_on = "poid", how = "left")
master = pd.merge(master, mxntr, left_on = "bkngid", right_on = "bkngid", how = "left")
master = pd.merge(master, wt_data, left_on = "bkngid", right_on = "bkngid", how = "left")

master["Tracking Container No"] = master["Container No"].fillna("NA")
master.loc[master["Tracking Container No"]=="NA","Tracking Container No"] = master["Import Container No"] 
#master["Tracking Container No"] = master["Tracking Container No"].replace("NA",master["Import Container No"])
master["Tracking Container No"] = master["Tracking Container No"].fillna("NA")
master["Tracking Container No"] = master["Tracking Container No"].replace("NA",master["WT Container(s)"])
master["Tracking Container No"] = master["Tracking Container No"].astype("str")
master["Tracking Container No"] = master["Tracking Container No"].str.replace(",","")
master["Tracking Container No"] = master["Tracking Container No"].apply(lambda x: x.split(' ')[0])
master["Tracking Container No"] = master["Tracking Container No"].str.replace("nan","NA")
master["Tracking Container No"] = master["Tracking Container No"].str.replace("SO#","NA")

del master["poid"]
del master["bkngid"]

master = pd.merge(master, gc_data, left_on = "Tracking Container No", right_on = "Tracking Container No", how = "left")
master = master.sort_values(by=["Division","Category","ID"], ascending=[True,True,True])

master["ETD"] = master["WT Port ETD"]
master["ETD"] = master["ETD"].fillna("NA")
master["ETD"] = master["ETD"].replace("NA",master["Origin Departure Actual Date"])
master["ETD"] = master["ETD"].fillna("NA")
master["ETD"] = master["ETD"].replace("NA",master["Origin Departure Planned Date (ETD)"])
master["ETD"] = master["ETD"].fillna("NA")
master["ETD"] = master["ETD"].replace("NA",master["Import ETD Port"])
master["ETD"] = master["ETD"].fillna("NA")
master["ETD"] = master["ETD"].replace("NA",master["AIMS ETD Port"])
master["ETD"] = master["ETD"].fillna("NA")
master["ETD"] = master["ETD"].replace("NA",master["PO Cancel"])

master["ETA"] = master["WT Port ETA"]
master["ETA"] = master["ETA"].fillna("NA")
master["ETA"] = master["ETA"].replace("NA",master["Destination Arrival Actual Date"])
master["ETA"] = master["ETA"].fillna("NA")
master["ETA"] = master["ETA"].replace("NA",master["Destination Arrival Planned Date (ETA)"])
master["ETA"] = master["ETA"].fillna("NA")
master["ETA"] = master["ETA"].replace("NA",master["Import ETA Port"])
master["ETA"] = master["ETA"].fillna("NA")
master["ETA"] = master["ETA"].replace("NA",master["AIMS ETA Port"])
master["ETA"] = master["ETA"].fillna("NA")

master["Main Customer"] = " "
master["Priority"] = " "
master["Remarks"] = " "

master = master.rename(index = str, columns = {"ETD":"ETD Port","ETA":"ETA Port","Vendor Port":"Delivery Port",
                                               "Tracking Container No":"Tracking Container(s)"})

master = master[["Style","Color","PO No","Qty","Ship to Location","AIMS Location Code","Main Customer",
                 "Tracking Container(s)","ETD Port","ETA Port","Division","Category","Type","Status","Priority",
                 "PO Date","PO Cancel","Shipment No","Document No","Container No","Casepack","CRTNs","CRTN CBM",
                 "Total CBM","Total CBF","AIMS ETD Port","AIMS ETA Port","Origin Port","Delivery Port","Booking No",
                 "Import Container No","Import ETD Port","Import ETA Port","Import ETA Whs","WT Container(s)",
                 "WT Port ETD","WT Port ETA","Carrier Name","Origin Departure Vessel Name","Origin Port","Destination Port",
                 "Shipment Status","Origin Departure Planned Date (ETD)","Origin Departure Actual Date",
                 "Destination Arrival Planned Date (ETA)","Destination Arrival Actual Date","Remarks"]]                        
                               
################################################### SAVE DATA TO EXCEL ###################################################

writer = pd.ExcelWriter(r"C:\Users\Joshua Kemperman\Enchante Living\Planning - Documents\25 - PO TRACKING\Inbound Shipment Report Working.xlsx", engine = "xlsxwriter")
master.to_excel(writer,"Shipping Master",index=False)
#wt_data.to_excel(writer,"WT Data",index=False)
writer.save()

#master.to_csv(r"C:\Users\andre\Enchante Living\Planning - Documents\05 - INVENTORY\Inventory Master_New.csv")
master.head(50)
