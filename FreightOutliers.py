# Authors: John Ayres
# Desc: identifies freight cost outliers through a normalized z-score. data is log normal.
# PARAM: Inventory Report.xlsx
# OUT: xlsx file cleanly listing PO's and relative information in their own cell
import pandas as pd
import numpy as np


################################################### PATHS AND READ #####################################################

invp = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\05 - INVENTORY\Inventory Report_July 25, 2022.xlsx'

inv = pd.read_excel(invp, engine='openpyxl', header=2)

threshold = 1.86 #Z-score threshold (2 is less conservative, 1.86 is traditional value used)

save_Loc = r'C:\Users\John Ayres\OneDrive - Enchante Living\Documents\39 Joint Project\Freight Outlier Report TODAY.xlsx'
############################################# Categorization and data type #############################################

inv["Category"] = inv["Category"].astype('category')
inv["Ocean Freight Cost"] = inv["Ocean Freight Cost"].astype(float)
grouped = inv.groupby(inv["Category"])
basi = grouped.get_group("Basin")
bacc = grouped.get_group("Bathroom Acc")
bfau = grouped.get_group("Bathroom Faucet")
chan = grouped.get_group("Chandelier")
disp = grouped.get_group("Display")
flam = grouped.get_group("Floor Lamp")
flus = grouped.get_group("Flushmount")
isli = grouped.get_group("Island Light")
kfac = grouped.get_group("Kitchen Faucet")
ksin = grouped.get_group("Kitchen Sink")
kwar = grouped.get_group("Kitchenware")
mirr = grouped.get_group("Mirror")
olig = grouped.get_group("Outdoor Light")
pend = grouped.get_group("Pendant")
repp = grouped.get_group("Replacement Parts")
scur = grouped.get_group("Shower Curtain")
sdor = grouped.get_group("Shower Door")
sfau = grouped.get_group("Shower Faucet")
shea = grouped.get_group("Showerhead")
tlam = grouped.get_group("Table Lamp")
toil = grouped.get_group("Toilet")
tlig = grouped.get_group("Track Lighting")
ucab = grouped.get_group("Under Cabinet")
vani = grouped.get_group("Vanity")
vlig = grouped.get_group("Vanity Light")
whar = grouped.get_group("Wall Hardware")
wsco = grouped.get_group("Wall Sconce")


############################################# Log Transformation and tests #############################################

basi['Ocean Freight Cost'] = np.log(basi['Ocean Freight Cost'])
bacc['Ocean Freight Cost'] = np.log(bacc['Ocean Freight Cost'])
bfau['Ocean Freight Cost'] = np.log(bfau['Ocean Freight Cost'])
chan['Ocean Freight Cost'] = np.log(chan['Ocean Freight Cost'])
disp['Ocean Freight Cost'] = np.log(disp['Ocean Freight Cost'])
flam['Ocean Freight Cost'] = np.log(flam['Ocean Freight Cost'])
flus['Ocean Freight Cost'] = np.log(flus['Ocean Freight Cost'])
isli['Ocean Freight Cost'] = np.log(isli['Ocean Freight Cost'])
kfac['Ocean Freight Cost'] = np.log(kfac['Ocean Freight Cost'])
ksin['Ocean Freight Cost'] = np.log(ksin['Ocean Freight Cost'])
kwar['Ocean Freight Cost'] = np.log(kwar['Ocean Freight Cost'])
mirr['Ocean Freight Cost'] = np.log(mirr['Ocean Freight Cost'])
olig['Ocean Freight Cost'] = np.log(olig['Ocean Freight Cost'])
pend['Ocean Freight Cost'] = np.log(pend['Ocean Freight Cost'])
repp['Ocean Freight Cost'] = np.log(repp['Ocean Freight Cost'])
scur['Ocean Freight Cost'] = np.log(scur['Ocean Freight Cost'])
sdor['Ocean Freight Cost'] = np.log(sdor['Ocean Freight Cost'])
sfau['Ocean Freight Cost'] = np.log(sfau['Ocean Freight Cost'])
shea['Ocean Freight Cost'] = np.log(shea['Ocean Freight Cost'])
tlam['Ocean Freight Cost'] = np.log(tlam['Ocean Freight Cost'])
toil['Ocean Freight Cost'] = np.log(toil['Ocean Freight Cost'])
tlig['Ocean Freight Cost'] = np.log(tlig['Ocean Freight Cost'])
ucab['Ocean Freight Cost'] = np.log(ucab['Ocean Freight Cost'])
vani['Ocean Freight Cost'] = np.log(vani['Ocean Freight Cost'])
vlig['Ocean Freight Cost'] = np.log(vlig['Ocean Freight Cost'])
whar['Ocean Freight Cost'] = np.log(whar['Ocean Freight Cost'])
wsco['Ocean Freight Cost'] = np.log(wsco['Ocean Freight Cost'])

dframes = [basi, bacc, bfau, chan, disp, flam, flus, isli, kfac, ksin, kwar, mirr, olig, pend, repp, scur, sdor, sfau,
           shea, tlam, toil, tlig, ucab, vani, vlig, whar, wsco]

basi.reset_index(drop=True, inplace=True)
bacc.reset_index(drop=True, inplace=True)
bfau.reset_index(drop=True, inplace=True)
chan.reset_index(drop=True, inplace=True)
disp.reset_index(drop=True, inplace=True)
flam.reset_index(drop=True, inplace=True)
flus.reset_index(drop=True, inplace=True)
isli.reset_index(drop=True, inplace=True)
kfac.reset_index(drop=True, inplace=True)
ksin.reset_index(drop=True, inplace=True)
kwar.reset_index(drop=True, inplace=True)
mirr.reset_index(drop=True, inplace=True)
olig.reset_index(drop=True, inplace=True)
pend.reset_index(drop=True, inplace=True)
repp.reset_index(drop=True, inplace=True)
scur.reset_index(drop=True, inplace=True)
sdor.reset_index(drop=True, inplace=True)
sfau.reset_index(drop=True, inplace=True)
shea.reset_index(drop=True, inplace=True)
tlam.reset_index(drop=True, inplace=True)
toil.reset_index(drop=True, inplace=True)
tlig.reset_index(drop=True, inplace=True)
ucab.reset_index(drop=True, inplace=True)
vani.reset_index(drop=True, inplace=True)
vlig.reset_index(drop=True, inplace=True)
whar.reset_index(drop=True, inplace=True)
wsco.reset_index(drop=True, inplace=True)


################################################## Calculating Fails ###################################################

def zscore(data):
    if len(data) == 0:
        return null
    else:
        mean = np.mean(data.loc[:, "Ocean Freight Cost"])
        std = np.std(data.loc[:, "Ocean Freight Cost"])
        if std == 0:
            data.loc[:, "Outlier"] = "False"
        else:
            count = 0
            for i in data["Ocean Freight Cost"]:
                z = (i - mean) / std
                data.loc[count, "Z-score"] = z
                if z > threshold:
                    data.loc[count, "Outlier"] = "True"
                else:
                    data.loc[count, "Outlier"] = "False"
                count = count + 1



zscore(basi)
zscore(bacc)
zscore(bfau)
zscore(chan)
zscore(disp)
zscore(flam)
zscore(flus)
zscore(isli)
zscore(kfac)
zscore(ksin)
zscore(kwar)
zscore(mirr)
zscore(olig)
zscore(pend)
zscore(repp)
zscore(scur)
zscore(sdor)
zscore(sfau)
zscore(shea)
zscore(tlam)
zscore(toil)
zscore(tlig)
zscore(ucab)
zscore(vani)
zscore(vlig)
zscore(whar)
zscore(wsco)


################################################## Creating Fails DF ###################################################
outlier = pd.DataFrame()
outlier = pd.concat([outlier, basi[basi["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, bacc[bacc["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, bfau[bfau["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, chan[chan["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, disp[disp["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, flam[flam["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, flus[flus["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, isli[isli["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, kfac[kfac["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, ksin[ksin["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, kwar[kwar["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, mirr[mirr["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, olig[olig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, pend[pend["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, repp[repp["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, scur[scur["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, sdor[sdor["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, sfau[sfau["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, shea[shea["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, tlam[tlam["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, toil[toil["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, tlig[tlig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, ucab[ucab["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, vani[vani["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, vlig[vlig["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, whar[whar["Outlier"] == "True"]], ignore_index=True)
outlier = pd.concat([outlier, wsco[wsco["Outlier"] == "True"]], ignore_index=True)


################################################ Inverse Transformation ################################################

outlier["Ocean Freight Cost"] = np.exp(outlier["Ocean Freight Cost"])


##################################################### FILE OUTPUTS #####################################################

fileName = pd.ExcelWriter(save_Loc, engine='xlsxwriter')

outlier.to_excel(fileName, sheet_name='Outliers', index=False)
inv.to_excel(fileName, sheet_name='Inventory Report', index=False)

fileName.save()


# So i dont have to type a bunch
# basi
# bacc
# bfau
# chan
# disp
# flam
# flus
# isli
# kfac
# ksin
# kwar
# mirr
# olig
# pend
# repp
# scur
# sdor
# sfau
# shea
# tlam
# toil
# tlig
# ucab
# vani
# vlig
# whar
# wsco
