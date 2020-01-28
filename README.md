# automation
## This is a simple tutorial for automation
For corporate financial analyst,  this is a great tool to help simplify all repeating tasks. 
Most of the analysts will receive weekly or monthly report in excel or pdf or csv or other formate.
Our goal is to spend most of the time to do actual analysis instead of organizing spreadsheet or physically input data.
The subject example is based on a hotel operation. Please see STR Mothly Report.py

![automation](https://user-images.githubusercontent.com/46503526/73231851-5d46f100-4179-11ea-99f6-700f147bff87.jpg)

## Step one 

- find pattern of files you received
- find critical info needed for analysis EG: current month Occupancy, ADR, RevPar
- find cell location number and assign each value to a variable 

![STAR](https://user-images.githubusercontent.com/46503526/73231891-7cde1980-4179-11ea-9f1a-7938b1956d19.PNG)

## Step two 
- open pgAdmin for Postgresql and create a database and a table 

![1](https://user-images.githubusercontent.com/46503526/73232389-faeef000-417a-11ea-8261-1da557182b48.PNG)
- open Query Tool and enter code to create table to host the data needed (F5).  
create table mstar(date date, Occ real, Occ_growth real, Occ_comp real, Occ_comp_growth real,MPI real,MPI_growth real)
![2](https://user-images.githubusercontent.com/46503526/73232568-92544300-417b-11ea-834d-c3bea7d96928.PNG)
- check the dataframe and make sure it fit for the data you want to scrape from files

## Step three 
- execute the python file to insert the value 
- conn.commit(); this is the actual command to make it effective
- everytime you run the python code will automated create a row of record in postgresql

## Step four
- use business intelligence to analyse and visualize your data and create dash-board
- the most famous BI is tableau which is costly
![3](https://user-images.githubusercontent.com/46503526/73233023-3e356880-4153-11ea-9d8c-c7544ccae94f.PNG)
- you can use powerbi from microsoft instead. It does not support Postgresql. You can download excel from Postgresql and upload into powerbi. 
![4](https://user-images.githubusercontent.com/46503526/73233123-8f455c80-4153-11ea-8928-4643ed0e45b4.PNG)

## other example
- Marriott Weekly Segmentation Report
- Marriott Weekly PACE Report
- M3 Monthly  P&L Report 
