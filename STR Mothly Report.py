import xlrd
import psycopg2
from datetime import datetime

path = r"C:\abc.xls"

book = xlrd.open_workbook(path)
sheet1 = book.sheet_by_name("Glance")

v1 = sheet1.cell(10, 6).value  # Occ
v2 = sheet1.cell(24, 6).value  # Occ_growth
v3 = sheet1.cell(10, 7).value  # Occ_comp
v4 = sheet1.cell(24, 7).value  # occ_comp_growth
v5 = sheet1.cell(10, 8).value  # MPI
v6 = sheet1.cell(24, 8).value  # MPI_growth
v7 = sheet1.cell(10, 11).value  # Adr
v8 = sheet1.cell(24, 11).value  # Adr_growth
v9 = sheet1.cell(10, 12).value  # Adr_comp
v10 = sheet1.cell(24, 12).value  # Adr_comp_growth
v11 = sheet1.cell(10, 13).value  # ARI
v12 = sheet1.cell(24, 13).value  # ARI_growth
v13 = sheet1.cell(10, 16).value  # Rev
v14 = sheet1.cell(24, 16).value  # Rev_growth
v15 = sheet1.cell(10, 17).value  # Rev_comps
v16 = sheet1.cell(24, 17).value  # Rev_comps_growth
v17 = sheet1.cell(10, 18).value  # RGI
v18 = sheet1.cell(24, 18).value  # RGI_growth
vda = sheet1.cell(5, 1)  # Date
vdb = vda.value
vd = datetime.strptime(vdb, '%B %Y').strftime('%Y-%m-%d')

sheet2 = book.sheet_by_name("Summary")

vs1 = sheet2.cell(8, 4).value  # Market_FM_Occ
vs2 = sheet2.cell(8, 5).value  # Market_FM_Occ_growth
vs3 = sheet2.cell(9, 4).value  # Market_Upmscale_Occ
vs4 = sheet2.cell(9, 5).value  # Market_Upmscale_Occ_growth
vs5 = sheet2.cell(10, 4).value  # Tract_FM_Occ
vs6 = sheet2.cell(10, 5).value  # Tract_FM_Occ_growth
vs7 = sheet2.cell(11, 4).value  # Tract_midscale_Occ
vs8 = sheet2.cell(11, 5).value  # Tract_midscale_Occ_growth

vs9 = sheet2.cell(18, 4).value  # Market_FM_Adr
vs10 = sheet2.cell(18, 5).value  # Market_FM_Adr_growth
vs11 = sheet2.cell(19, 4).value  # Market_Upmscale_Adr
vs12 = sheet2.cell(19, 5).value  # Market_Upmscale_Adr_Growth
vs13 = sheet2.cell(20, 4).value  # Tract_FM_Adr
vs14 = sheet2.cell(20, 5).value  # Tract_FM_Adr_growth
vs15 = sheet2.cell(21, 4).value  # Tract_midscale_Adr
vs16 = sheet2.cell(21, 5).value  # Tract_midscale_Adr_growth

vs17 = sheet2.cell(28, 4).value  # Market_FM_Rev
vs18 = sheet2.cell(28, 5).value  # Market_FM_Rev_growth
vs19 = sheet2.cell(29, 4).value  # Market_Upmscale_Rev
vs20 = sheet2.cell(29, 5).value  # Market_Upmscale_Rev_Growth
vs21 = sheet2.cell(30, 4).value  # Tract_FM_Rev
vs22 = sheet2.cell(30, 5).value  # Tract_FM_Rev_growth
vs23 = sheet2.cell(31, 4).value  # Tract_midscale_Rev
vs24 = sheet2.cell(31, 5).value  # Tract_midscale_Rev_growth

vs25 = sheet2.cell(7, 13).value  # Market_FM_supply
vs26 = sheet2.cell(8, 13).value  # Market_Upmscale_supply
vs27 = sheet2.cell(9, 13).value  # Tract_FM_supply
vs28 = sheet2.cell(10, 13).value  # Tract_midscale_supply

vs29 = sheet2.cell(18, 13).value  # Market_FM_demand       $ Occupancy comparison
vs30 = sheet2.cell(19, 13).value  # Market_Upmscale_demand
vs31 = sheet2.cell(20, 13).value  # Tract_FM_demand
vs32 = sheet2.cell(21, 13).value  # Tract_midscale_demand

vs33 = sheet2.cell(28, 13).value  # Market_FM_revenue      # Revpar comparison
vs34 = sheet2.cell(29, 13).value  # Market_Upmscale_revenue
vs35 = sheet2.cell(30, 13).value  # Tract_FM_revenue
vs36 = sheet2.cell(31, 13).value  # Tract_midscale_revenue

vs37 = sheet2.cell(37, 4).value  # Market_FM_properties
vs38 = sheet2.cell(37, 6).value  # Market_FM_rooms
vs39 = sheet2.cell(38, 4).value  # Market_Upmscale_properties
vs40 = sheet2.cell(38, 6).value  # Market_Upmscale_rooms
vs41 = sheet2.cell(39, 4).value  # Tract_FM_properties
vs42 = sheet2.cell(39, 6).value  # Tract_FM_rooms
vs43 = sheet2.cell(40, 4).value  # Tract_midscale_properties
vs44 = sheet2.cell(40, 6).value  # Tract_midscale_rooms

vs45 = sheet2.cell(38, 13).value  # Construct_property
vs46 = sheet2.cell(38, 14).value  # Construct_room
vs47 = sheet2.cell(38, 16).value  # Planning_property
vs48 = sheet2.cell(38, 18).value  # Planning_room

# Input data into sql
conn = psycopg2.connect("host='localhost' port='5432' dbname='hotel' user='postgres' password='...'")
cur = conn.cursor()
# create table mstar(date date, Occ real, Occ_growth real, Occ_comp real, Occ_comp_growth real,
# MPI real,MPI_growth real, Adr real,Adr_growth real,Adr_comp real,Adr_comp_growth real,
# ARI real, ARI_growth real,Rev real, Rev_growth real, Rev_comp real,Rev_comp_growth real,
# RGI real, RGI_growth real,
# Market_Col_Occ real,Market_Col_Occ_growth real,Market_Upscale_Occ real,Market_Upscale_Occ_growth real,
# Tract_Col_Occ real,Tract_Col_Occ_growth real,Tract_upscale_Occ real,Tract_upscale_Occ_growth real,
# Market_Col_Adr real,Market_Col_Adr_growth real,Market_Upscale_Adr real,Market_Upscale_Adr_growth real,
# Tract_Col_Adr real,Tract_Col_Adr_growth real,Tract_upscale_Adr real,Tract_upscale_Adr_growth real,
# Market_Col_Rev real,Market_Col_Rev_growth real,Market_Upscale_Rev real,Market_Upscale_Rev_growth real,
# Tract_Col_Rev real,Tract_Col_Rev_growth real,Tract_upscale_Rev real,Tract_upscale_Rev_growth real,
# Market_Col_supply real, Market_Upscale_supply real, Tract_Col_supply real, Tract_upscale_supply real,
# Market_Col_demand real, Market_Upscale_demand real, Tract_Col_demand real, Tract_upscale_demand real,
# Market_Col_revenue real, Market_Upscale_revenue real, Tract_Col_revenue real, Tract_upscale_revenue real,
# Market_Col_properties real,Market_Col_rooms real,Market_Upscale_properties real,Market_Upscale_rooms real,
# Tract_Col_properties real,Tract_Col_rooms real,Tract_upscale_properties real,Tract_upscale_rooms real,
# Construct_property real,Construct_room real,Planning_property real,Planning_room real);
cur.execute("insert into Mstar values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,"
            "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (vd, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18,
             vs1, vs2, vs3, vs4, vs5, vs6, vs7, vs8, vs9, vs10, vs11, vs12, vs13, vs14, vs15, vs16, vs17, vs18, vs19,
             vs20, vs21, vs22, vs23, vs24,
             vs25, vs26, vs27, vs28, vs29, vs30, vs31, vs32, vs33, vs34, vs35, vs36, vs37, vs38, vs39, vs40, vs41, vs42,
             vs43, vs44,
             vs45, vs46, vs47, vs48
             ))

conn.commit()
