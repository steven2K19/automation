import psycopg2
from datetime import datetime
import pandas as pd

path = r"C:\abc.xlsx"
# Pick up Report
df = pd.read_excel(path, sheetname=['MTD', 'PACE', 'Pickup'], header=None)
sheet2 = df['PACE']
vd1 = datetime(2019, 5, 8)
vd = vd1.strftime('%Y-%m-%d')

# Current month Projection
# CMP1_CY_RN, CMP1_CY_ADR, CMP1_CY_Rev,CMP1_LY_RN, CMP1_LY_ADR, CMP1_LY_Rev,CMP1_Var_RN,CMP1_Var_ADR,CMP1_Var_Rev,
# CMP2_CY_RN, CMP2_CY_ADR, CMP2_CY_Rev,CMP2_LY_RN, CMP2_LY_ADR, CMP2_LY_Rev,CMP2_Var_RN,CMP2_Var_ADR,CMP2_Var_Rev,
v1 = sheet2.loc[13, 1]
v2 = sheet2.loc[13, 2]
v3 = sheet2.loc[13, 3]
v4 = sheet2.loc[13, 7]
v5 = sheet2.loc[13, 8]
v6 = sheet2.loc[13, 9]
v7 = sheet2.loc[13, 10]
v8 = sheet2.loc[13, 11]
v9 = sheet2.loc[13, 12]
v10 = sheet2.loc[14, 1]
v11 = sheet2.loc[14, 2]
v12 = sheet2.loc[14, 3]
v13 = sheet2.loc[14, 7]
v14 = sheet2.loc[14, 8]
v15 = sheet2.loc[14, 9]
v16 = sheet2.loc[14, 10]
v17 = sheet2.loc[14, 11]
v18 = sheet2.loc[14, 12]

# Next month Projection
# NMP1_CY_RN, NMP1_CY_ADR, NMP1_CY_Rev,NMP1_LY_RN, NMP1_LY_ADR, NMP1_LY_Rev,NMP1_Var_RN,NMP1_Var_ADR,NMP1_Var_Rev,
# NMP2_CY_RN, NMP2_CY_ADR, NMP2_CY_Rev,NMP2_LY_RN, NMP2_LY_ADR, NMP2_LY_Rev,NMP2_Var_RN,NMP2_Var_ADR,NMP2_Var_Rev,
v19 = sheet2.loc[55, 1]
v20 = sheet2.loc[55, 2]
v21 = sheet2.loc[55, 3]
v22 = sheet2.loc[55, 7]
v23 = sheet2.loc[55, 8]
v24 = sheet2.loc[55, 9]
v25 = sheet2.loc[55, 10]
v26 = sheet2.loc[55, 11]
v27 = sheet2.loc[55, 12]
v28 = sheet2.loc[56, 1]
v29 = sheet2.loc[56, 2]
v30 = sheet2.loc[56, 3]
v31 = sheet2.loc[56, 7]
v32 = sheet2.loc[56, 8]
v33 = sheet2.loc[56, 9]
v34 = sheet2.loc[56, 10]
v35 = sheet2.loc[56, 11]
v36 = sheet2.loc[56, 12]

# The month of Next month Projection
# NNMP1_CY_RN,NNMP1_CY_ADR,NNMP1_CY_Rev,NNMP1_LY_RN,NNMP1_LY_ADR,NNMP1_LY_Rev,NNMP1_Var_RN,NNMP1_Var_ADR,NNMP1_Var_Rev,
# NNMP2_CY_RN,NNMP2_CY_ADR,NNMP2_CY_Rev,NNMP2_LY_RN,NNMP2_LY_ADR,NNMP2_LY_Rev,NNMP2_Var_RN,NNMP2_Var_ADR,NNMP2_Var_Rev,
v37 = sheet2.loc[97, 1]
v38 = sheet2.loc[97, 2]
v39 = sheet2.loc[97, 3]
v40 = sheet2.loc[97, 7]
v41 = sheet2.loc[97, 8]
v42 = sheet2.loc[97, 9]
v43 = sheet2.loc[97, 10]
v44 = sheet2.loc[97, 11]
v45 = sheet2.loc[97, 12]
v46 = sheet2.loc[98, 1]
v47 = sheet2.loc[98, 2]
v48 = sheet2.loc[98, 3]
v49 = sheet2.loc[98, 7]
v50 = sheet2.loc[98, 8]
v51 = sheet2.loc[98, 9]
v52 = sheet2.loc[98, 10]
v53 = sheet2.loc[98, 11]
v54 = sheet2.loc[98, 12]

# Input data into sql
conn = psycopg2.connect("host='localhost' port='5432' dbname='CYCO' user='postgres' password='Qweasd'")
cur = conn.cursor()
# create table PACE(date date,
# CMP1_CY_RN real, CMP1_CY_ADR real, CMP1_CY_Rev real,CMP1_LY_RN real, CMP1_LY_ADR real, CMP1_LY_Rev real,
# CMP1_Var_RN real,CMP1_Var_ADR real,CMP1_Var_Rev real,
# CMP2_CY_RN real, CMP2_CY_ADR real, CMP2_CY_Rev real,CMP2_LY_RN real, CMP2_LY_ADR real, CMP2_LY_Rev real,
# CMP2_Var_RN real,CMP2_Var_ADR real,CMP2_Var_Rev real,
# NMP1_CY_RN real, NMP1_CY_ADR real, NMP1_CY_Rev real,NMP1_LY_RN real, NMP1_LY_ADR real, NMP1_LY_Rev real,
# NMP1_Var_RN real,NMP1_Var_ADR real,NMP1_Var_Rev real,
# NMP2_CY_RN real, NMP2_CY_ADR real, NMP2_CY_Rev real,NMP2_LY_RN real, NMP2_LY_ADR real, NMP2_LY_Rev real,
# NMP2_Var_RN real,NMP2_Var_ADR real,NMP2_Var_Rev real,
# NNMP1_CY_RN real,NNMP1_CY_ADR real,NNMP1_CY_Rev real,NNMP1_LY_RN real,NNMP1_LY_ADR real,NNMP1_LY_Rev real,
# NNMP1_Var_RN real,NNMP1_Var_ADR real,NNMP1_Var_Rev real,
# NNMP2_CY_RN real,NNMP2_CY_ADR real,NNMP2_CY_Rev real,NNMP2_LY_RN real,NNMP2_LY_ADR real,NNMP2_LY_Rev real,
# NNMP2_Var_RN real,NNMP2_Var_ADR real,NNMP2_Var_Rev real);

cur.execute("insert into PACE values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (vd, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18,
             v19, v20, v21, v22, v23, v24, v25, v26, v27, v28, v29, v30, v31, v32,
             v33, v34, v35, v36, v37, v38, v39, v40, v41, v42, v43, v44, v45, v46, v47, v48, v49, v50,
             v51, v52, v53, v54))

conn.commit()
