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
- open Query Tool and enter code to create table to host the data needed
create table mstar(date date, Occ real, Occ_growth real, Occ_comp real, Occ_comp_growth real,MPI real,MPI_growth real)
![2](https://user-images.githubusercontent.com/46503526/73232568-92544300-417b-11ea-834d-c3bea7d96928.PNG)
- check the dataframe and make sure it fit for the data you want to scrape from files

## Step three 
- financials for recent 3 years. (income statement, balance sheet, cash flow statement)
- calendar year end financials prorated by months
- Trailing twelve months financial by aggregate recent 4 quarters
- Basic share and price for each fiscal year end and recent fiscal quarter end
- It takes about 10 minutes to create a dataframe for 100 stocks. 
