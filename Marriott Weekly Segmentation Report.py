import psycopg2
from datetime import datetime
import pandas as pd

path = r"abc.xlsx"
# Pick up Report
df = pd.read_excel(path, sheetname=['MTD', 'PACE', 'Pickup'], header=None)
sheet1 = df['MTD']
sheet2 = df['PACE']
sheet3 = df['Pickup']
vd1 = datetime(2019, 5, 8)
vd = vd1.strftime('%Y-%m-%d')
# Total transient OTB (CM_Total_CY_OTB, CM_Total_CY_Pickup, CM_Total_LY_OTB, CM_Total_LY_Pickup)
v1 = sheet3.loc[33, 13]
v2 = sheet3.loc[33, 14]
v3 = sheet3.loc[33, 15]
v4 = sheet3.loc[33, 16]
# STD Retail (CM_STDRetail_CY_OTB, CM_STDRetail_CY_Pickup, CM_STDRetail_LY_OTB, CM_STDRetail_LY_Pickup)
v5 = sheet3.loc[33, 23]
v6 = sheet3.loc[33, 24]
v7 = sheet3.loc[33, 25]
v8 = sheet3.loc[33, 26]
# LNR (CM_LNR_CY_OTB, CM_LNR_CY_Pickup, CM_LNR_LY_OTB, CM_LNR_LY_Pickup)
v9 = sheet3.loc[33, 28]
v10 = sheet3.loc[33, 29]
v11 = sheet3.loc[33, 30]
v12 = sheet3.loc[33, 31]
# Disc Retail (CM_DiscRetail_CY_OTB, CM_DiscRetail_CY_Pickup, CM_DiscRetail_LY_OTB, CM_DiscRetail_LY_Pickup)
v13 = sheet3.loc[33, 33]
v14 = sheet3.loc[33, 34]
v15 = sheet3.loc[33, 35]
v16 = sheet3.loc[33, 36]
# Disc Fix (CM_DiscFix_CY_OTB, CM_DiscFix_CY_Pickup, CM_DiscFix_LY_OTB, CM_DiscFix_LY_Pickup)
v17 = sheet3.loc[33, 38]
v18 = sheet3.loc[33, 39]
v19 = sheet3.loc[33, 40]
v20 = sheet3.loc[33, 41]
# Gov (CM_Gov_CY_OTB, CM_Gov_CY_Pickup, CM_Gov_LY_OTB, CM_Gov_LY_Pickup)
v21 = sheet3.loc[33, 43]
v22 = sheet3.loc[33, 44]
v23 = sheet3.loc[33, 45]
v24 = sheet3.loc[33, 46]
# Reward (CM_Reward_CY_OTB, CM_Reward_CY_Pickup, CM_Reward_LY_OTB, CM_Reward_LY_Pickup)
v25 = sheet3.loc[33, 48]
v26 = sheet3.loc[33, 49]
v27 = sheet3.loc[33, 50]
v28 = sheet3.loc[33, 51]
# Policy (CM_Policy_CY_OTB, CM_Policy_CY_Pickup, CM_Policy_LY_OTB, CM_Policy_LY_Pickup)
v29 = sheet3.loc[33, 53]
v30 = sheet3.loc[33, 54]
v31 = sheet3.loc[33, 55]
v32 = sheet3.loc[33, 56]

# Total transient OTB (CM_Total_CY_OTB, CM_Total_CY_Pickup, CM_Total_LY_OTB, CM_Total_LY_Pickup)
v33 = sheet3.loc[72, 13]
v34 = sheet3.loc[72, 14]
v35 = sheet3.loc[72, 15]
v36 = sheet3.loc[72, 16]
# STD Retail (CM_STDRetail_CY_OTB, CM_STDRetail_CY_Pickup, CM_STDRetail_LY_OTB, CM_STDRetail_LY_Pickup)
v37 = sheet3.loc[72, 23]
v38 = sheet3.loc[72, 24]
v39 = sheet3.loc[72, 25]
v40 = sheet3.loc[72, 26]
# LNR (CM_LNR_CY_OTB, CM_LNR_CY_Pickup, CM_LNR_LY_OTB, CM_LNR_LY_Pickup)
v41 = sheet3.loc[72, 28]
v42 = sheet3.loc[72, 29]
v43 = sheet3.loc[72, 30]
v44 = sheet3.loc[72, 31]
# Disc Retail (CM_DiscRetail_CY_OTB, CM_DiscRetail_CY_Pickup, CM_DiscRetail_LY_OTB, CM_DiscRetail_LY_Pickup)
v45 = sheet3.loc[72, 33]
v46 = sheet3.loc[72, 34]
v47 = sheet3.loc[72, 35]
v48 = sheet3.loc[72, 36]
# Disc Fix (CM_DiscFix_CY_OTB, CM_DiscFix_CY_Pickup, CM_DiscFix_LY_OTB, CM_DiscFix_LY_Pickup)
v49 = sheet3.loc[72, 38]
v50 = sheet3.loc[72, 39]
v51 = sheet3.loc[72, 40]
v52 = sheet3.loc[72, 41]
# Gov (CM_Gov_CY_OTB, CM_Gov_CY_Pickup, CM_Gov_LY_OTB, CM_Gov_LY_Pickup)
v53 = sheet3.loc[72, 43]
v54 = sheet3.loc[72, 44]
v55 = sheet3.loc[72, 45]
v56 = sheet3.loc[72, 46]
# Reward (CM_Reward_CY_OTB, CM_Reward_CY_Pickup, CM_Reward_LY_OTB, CM_Reward_LY_Pickup)
v57 = sheet3.loc[72, 48]
v58 = sheet3.loc[72, 49]
v59 = sheet3.loc[72, 50]
v60 = sheet3.loc[72, 51]
# Policy (CM_Policy_CY_OTB, CM_Policy_CY_Pickup, CM_Policy_LY_OTB, CM_Policy_LY_Pickup)
v61 = sheet3.loc[72, 53]
v62 = sheet3.loc[72, 54]
v63 = sheet3.loc[72, 55]
v64 = sheet3.loc[72, 56]

# Input data into sql
conn = psycopg2.connect("host='localhost' port='5432' dbname='CYCO' user='postgres' password='xxx'")
cur = conn.cursor()
# create table pickup(date date,
# CM_Total_CY_OTB real, CM_Total_CY_Pickup real, CM_Total_LY_OTB real, CM_Total_LY_Pickup real,
# CM_STDRetail_CY_OTB real, CM_STDRetail_CY_Pickup real, CM_STDRetail_LY_OTB real, CM_STDRetail_LY_Pickup real,
# CM_LNR_CY_OTB real, CM_LNR_CY_Pickup real, CM_LNR_LY_OTB real, CM_LNR_LY_Pickup real,
# CM_DiscRetail_CY_OTB real, CM_DiscRetail_CY_Pickup real, CM_DiscRetail_LY_OTB real, CM_DiscRetail_LY_Pickup real,
# CM_DiscFix_CY_OTB real, CM_DiscFix_CY_Pickup real, CM_DiscFix_LY_OTB real, CM_DiscFix_LY_Pickup real,
# CM_Gov_CY_OTB real, CM_Gov_CY_Pickup real, CM_Gov_LY_OTB real, CM_Gov_LY_Pickup real,
# CM_Reward_CY_OTB real, CM_Reward_CY_Pickup real, CM_Reward_LY_OTB real, CM_Reward_LY_Pickup real,
# CM_Policy_CY_OTB real, CM_Policy_CY_Pickup real, CM_Policy_LY_OTB real, CM_Policy_LY_Pickup real,
# NM_Total_CY_OTB real, NM_Total_CY_Pickup real, NM_Total_LY_OTB real, NM_Total_LY_Pickup real,
# NM_STDRetail_CY_OTB real, NM_STDRetail_CY_Pickup real, NM_STDRetail_LY_OTB real, NM_STDRetail_LY_Pickup real,
# NM_LNR_CY_OTB real, NM_LNR_CY_Pickup real, NM_LNR_LY_OTB real, NM_LNR_LY_Pickup real,
# NM_DiscRetail_CY_OTB real, NM_DiscRetail_CY_Pickup real, NM_DiscRetail_LY_OTB real, NM_DiscRetail_LY_Pickup real,
# NM_DiscFix_CY_OTB real, NM_DiscFix_CY_Pickup real, NM_DiscFix_LY_OTB real, NM_DiscFix_LY_Pickup real,
# NM_Gov_CY_OTB real, NM_Gov_CY_Pickup real, NM_Gov_LY_OTB real, NM_Gov_LY_Pickup real,
# NM_Reward_CY_OTB real, NM_Reward_CY_Pickup real, NM_Reward_LY_OTB real, NM_Reward_LY_Pickup real,
# NM_Policy_CY_OTB real, NM_Policy_CY_Pickup real, NM_Policy_LY_OTB real, NM_Policy_LY_Pickup real);
cur.execute("insert into pickup values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (vd, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18,
             v19, v20, v21, v22, v23, v24, v25, v26, v27, v28, v29, v30, v31, v32,
             v33, v34, v35, v36, v37, v38, v39, v40, v41, v42, v43, v44, v45, v46, v47, v48, v49, v50,
             v51, v52, v53, v54, v55, v56, v57, v58, v59, v60, v61, v62, v63, v64))

conn.commit()
