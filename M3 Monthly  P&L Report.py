import xlrd
import psycopg2
from datetime import datetime

path = r"C:\abc.xlsx"
book = xlrd.open_workbook(path)
sheet1 = book.sheet_by_name("hotel")

v1 = int(sheet1.cell(3, 1).value)  # Room Revenue
v2 = int(sheet1.cell(4, 1).value)  # F&B Revenue
v3 = int(sheet1.cell(5, 1).value)  # Other Revenue
v4 = int(sheet1.cell(6, 1).value)  # Total Revenue
v5 = int(sheet1.cell(7, 1).value)  # Room Expense
v6 = int(sheet1.cell(8, 1).value)  # F&B Expense
v7 = int(sheet1.cell(9, 1).value)  # Other Expense
v8 = int(sheet1.cell(10, 1).value)  # IT Expense
v9 = int(sheet1.cell(11, 1).value)  # Department expense
v10 = int(sheet1.cell(12, 1).value)  # Department profit
v11 = int(sheet1.cell(13, 1).value)  # A&G Expense
v12 = int(sheet1.cell(14, 1).value)  # Sales Expense
v13 = int(sheet1.cell(15, 1).value)  # Maintain Expense
v14 = int(sheet1.cell(16, 1).value)  # Utility Expense
v15 = int(sheet1.cell(17, 1).value)  # SG&A Expense
v16 = int(sheet1.cell(18, 1).value)  # House Profit
v17 = int(sheet1.cell(19, 1).value)  # Management Fee
v18 = int(sheet1.cell(20, 1).value)  # Franchise Fee
v19 = int(sheet1.cell(21, 1).value)  # Fix Expense
v20 = int(sheet1.cell(23, 1).value)  # Non-operating Expense
v21 = int(sheet1.cell(24, 1).value)  # NOI
v22 = int(sheet1.cell(27, 1).value)  # Interest expense
v23 = int(sheet1.cell(28, 1).value)  # Pretax Profit


va = int(sheet1.cell(2, 1).value)
vd = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + va - 2).strftime('%Y-%m-%d')

# Input data into sql
conn = psycopg2.connect("host='localhost' port='5432' dbname='hotel' user='postgres' password='...'")
cur = conn.cursor()

# First need to Create dataframe in PostgreSQL
# create table pl(date date,
# Room_Revenue integer,FB_Revenue integer,Other_Revenue integer,Total_Revenue integer,
# Room_Expense integer,FB_Expense integer,Other_Expense integer,IT_Expense integer,
# Department_Expense integer,Department_profit integer,
# AG_Expense integer,Sale_Expense integer,Maintain_Expense integer,Utility_Expense integer,
# SGA_Expense integer,House_Profit integer,
# Management_Fee integer,Franchise_Fee integer,Fix_Expense integer,NonOp_Expense integer,NOI integer,
# Interest_expense integer,Pretax_Profit integer);
cur.execute("insert into pl values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (vd, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19, v20, v21, v22,
             v23))
conn.commit()
