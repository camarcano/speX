# -*- coding: utf-8 -*-
"""
Created on Sat Sep 25 14:42:30 2021

@author: Marbalza2
"""

import requests
import datetime
from datetime import date, timedelta
from bs4 import BeautifulSoup
import pandas as pd
import lxml
import openpyxl



def parse_array_from_fangraphs_html(start_date,end_date):
    """
    Take a HTML stats page from fangraphs and parse it out to a dataframe.
    """
    # parse input
    PITCHERS_URL = "https://www.fangraphs.com/leaders.aspx?pos=all&stats=pit&lg=all&qual=0&type=c%2C13%2C7%2C8%2C120%2C121%2C331%2C105%2C111%2C24%2C19%2C14%2C329%2C324%2C45%2C122%2C6%2C42%2C43%2C328%2C330%2C322%2C323%2C326%2C332&season=2021&month=1000&season1=2015&ind=0&team=&rost=&age=&filter=&players=&startdate={}&enddate={}&page=1_2000".format(start_date, end_date)
    PITCHERS_URL = "https://www.fangraphs.com/leaders.aspx?pos=all&stats=pit&lg=all&qual=0&type=c,13,7,8,120,121,331,105,111,24,19,14,329,324,45,122,6,42,43,328,330,322,323,326,332,31,30,29&season=2021&month=1000&season1=2015&ind=0&team=&rost=&age=&filter=&players=&startdate={}&enddate={}&page=1_2000".format(start_date, end_date)

    # request the data
    pitchers_html = requests.get(PITCHERS_URL).text
    soup = BeautifulSoup(pitchers_html, "lxml")
    table = soup.find("table", {"class": "rgMasterTable"})
    
    # get headers
    headers_html = table.find("thead").find_all("th")
    headers = []
    for header in headers_html:
        headers.append(header.text)

    # get rows
    rows = []
    rows_html = table.find("tbody").find_all("tr")
    for row in rows_html:
        row_data = []
        for cell in row.find_all("td"):
            row_data.append(cell.text)
        rows.append(row_data)
    
    return pd.DataFrame(rows, columns = headers)

def calc_speX(df, IP_limit):
    work_speX = df[['Name','Team', 'IP', 'G', 'GS','K%','BB%','SO','BB','TBF',
                    'Events','Barrels','CSW%','O-Swing%','Zone%',
                    'Pitches', 'Balls', 'Strikes']].copy()
    
    allowed_chars = set('0123456789.%')
    work_speX = work_speX[work_speX['Zone%'].apply(lambda x: set(x).issubset(allowed_chars))]
    
   
    work_speX.insert(loc=5, column='K-BB%', value=['' for i in range(work_speX.shape[0])])
    work_speX.insert(loc=6, column='pCRA', value=['' for i in range(work_speX.shape[0])])
    work_speX.insert(loc=17, column='O-Sw%+Z%', value=['' for i in range(work_speX.shape[0])])
    work_speX.insert(loc=18, column='speX', value=['' for i in range(work_speX.shape[0])])

    work_speX['K%'] = work_speX['K%'].map(lambda x: x.rstrip('%'))
    work_speX['BB%'] = work_speX['BB%'].map(lambda x: x.rstrip('%'))
    work_speX['CSW%'] = work_speX['CSW%'].map(lambda x: x.rstrip('%'))
    work_speX['O-Swing%'] = work_speX['O-Swing%'].map(lambda x: x.rstrip('%'))
    work_speX['Zone%'] = work_speX['Zone%'].map(lambda x: x.rstrip('%'))
    work_speX = work_speX.replace(r'^\s*$', 0, regex=True)    

    
    work_speX['K%'] = work_speX['K%'].astype(float)
    work_speX['BB%'] = work_speX['BB%'].astype(float)
    work_speX['CSW%'] = work_speX['CSW%'].astype(float)
    work_speX['O-Swing%'] = work_speX['O-Swing%'].astype(float)
    work_speX['SO'] = work_speX['SO'].astype(float)
    work_speX['TBF'] = work_speX['TBF'].astype(float)
    work_speX['BB'] = work_speX['BB'].astype(float)
    work_speX['Barrels'] = work_speX['Barrels'].astype(float)
    work_speX['Events'] = work_speX['Events'].astype(float)
    work_speX['IP'] = work_speX['IP'].astype(float)
    work_speX['Pitches'] = work_speX['Pitches'].astype(float)
    work_speX['Balls'] = work_speX['Balls'].astype(float)
    work_speX['Zone%'] = work_speX['Zone%'].astype(float)
    work_speX['Zone%'] = work_speX['Zone%'].fillna(0)


    

    work_speX['K-BB%'] = work_speX['K%'] - work_speX['BB%']
    work_speX['CSW-BB%'] = work_speX['CSW%'] - work_speX['BB%']
    work_speX['CSW-B%'] = work_speX['CSW%'] - (work_speX['Balls']/work_speX['Pitches'])*100
    work_speX['CSW-W%'] = work_speX['CSW%'] - (work_speX['BB']/work_speX['Pitches'])*100
    work_speX['O-Sw%+Z%'] = work_speX['O-Swing%'] + work_speX['Zone%']
    work_speX['pCRA'] = (-8.89 * ((work_speX['SO'] + 10.66) / (work_speX['TBF'] + 10.66 + 37.53)) + 8.03 * ((work_speX['BB'] + 9.58) /(work_speX['TBF'] + 9.58 + 116.2)) + 5.54 * ((work_speX['Barrels'] + 13.33) / (work_speX['Events'] + 13.33 + 191.5)) + 106 * ((work_speX['Barrels'] + 13.33) / (work_speX['Events'] + 13.33 + 191.5)) * ((work_speX['Barrels'] + 13.33) / (work_speX['Events'] + 13.33 + 191.5))) + 4.55
    work_speX['speX'] = (11 * work_speX['K-BB%'] + 165 * (10-work_speX['pCRA']) + 7 * work_speX['CSW%'] + work_speX['O-Sw%+Z%']) * 0.049

    speX = work_speX.drop(['K%', 'BB%', 'SO', 'BB', 'TBF', 'Events', 'Barrels', 'Strikes'], axis = 1)


    speX['K-BB%'] = speX['K-BB%'].round(decimals=2)
    speX['pCRA'] = speX['pCRA'].round(decimals=2)
    speX['CSW%'] = speX['CSW%'].round(decimals=2)
    speX['O-Swing%'] = speX['O-Swing%'].round(decimals=2)
    speX['Zone%'] = speX['Zone%'].round(decimals=2)
    speX['O-Sw%+Z%'] = speX['O-Sw%+Z%'].round(decimals=2)
    speX['speX'] = speX['speX'].round(decimals=2)


    speX = speX.sort_values(by=['speX'], ascending=False)

    speX = speX[speX['IP'] >= IP_limit]
    speX = speX.sort_values(by=['speX'], ascending=False)
    
    return speX

sdate = '2023-03-30'
enddate = '2023-04-10'
IP = 0 
date_format = "%Y-%m-%d"
start_date = datetime.datetime.strptime(sdate, date_format)
end_date = datetime.datetime.strptime(enddate, date_format)

daily = input("Do you want the daily values? ")
writer = pd.ExcelWriter('speX-daily.xlsx', engine='openpyxl') 

if daily.lower()=="y":
    for single_date in pd.date_range(start=start_date, end=end_date):
        date_str = single_date.strftime(date_format)
        speX = parse_array_from_fangraphs_html(date_str, date_str)
        result = calc_speX(speX, IP)
        result.to_excel(writer, sheet_name=date_str)





#date.today() - timedelta(1)
#enddate = enddate.strftime("%Y-%m-%d")


speX = parse_array_from_fangraphs_html(sdate, enddate)
speX_done = calc_speX(speX, IP)

speX_done.to_excel(writer, sheet_name='full')

#sdate = '2022-01-01'
#enddate = '2022-12-01'

sdate = date.today() - timedelta(16)
sdate = sdate.strftime("%Y-%m-%d")
IP = 0 

speX = parse_array_from_fangraphs_html(sdate, enddate)
speX_done = calc_speX(speX, IP)

speX_done.to_excel(writer, sheet_name='15 days')

sdate = date.today() - timedelta(31)
sdate = sdate.strftime("%Y-%m-%d")
IP = 0 

speX = parse_array_from_fangraphs_html(sdate, enddate)
speX_done = calc_speX(speX, IP)

speX_done.to_excel(writer, sheet_name='30 days')

sdate = date.today() - timedelta(46)
sdate = sdate.strftime("%Y-%m-%d")
IP = 0 

speX = parse_array_from_fangraphs_html(sdate, enddate)
speX_done = calc_speX(speX, IP)

speX_done.to_excel(writer, sheet_name='45 days')

writer.save()

writer.close()

# Open the Excel file
workbook = openpyxl.load_workbook('speX-daily.xlsx')

# Get the worksheet you want to move
worksheet = workbook['45 days']

# Get the current sheet index of the worksheet
sheet_index = workbook.index(worksheet)

# Move the worksheet to the beginning
workbook.move_sheet(worksheet, offset=-sheet_index)

# Get the worksheet you want to move
worksheet = workbook['30 days']

# Get the current sheet index of the worksheet
sheet_index = workbook.index(worksheet)

# Move the worksheet to the beginning
workbook.move_sheet(worksheet, offset=-sheet_index)

# Get the worksheet you want to move
worksheet = workbook['15 days']

# Get the current sheet index of the worksheet
sheet_index = workbook.index(worksheet)

# Move the worksheet to the beginning
workbook.move_sheet(worksheet, offset=-sheet_index)

# Get the worksheet you want to move
worksheet = workbook['full']

# Get the current sheet index of the worksheet
sheet_index = workbook.index(worksheet)

# Move the worksheet to the beginning
workbook.move_sheet(worksheet, offset=-sheet_index)

# Iterate through each worksheet in the workbook
for worksheet in workbook.worksheets:
# Auto-fit all columns in the worksheet
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width


# Save the changes to the Excel file
workbook.save('speX-daily.xlsx')

workbook.close()