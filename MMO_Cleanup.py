# Python 3.7.2
from datetime import datetime

from numpy import NaN
from pandas import set_option, read_excel, ExcelWriter

# Globals
now = datetime.now()
year, weeknum, dow = now.isocalendar()  # Current year, weeknumber, day of week
month = str(now.strftime("%m"))
# -------------------------
# -----1. Import File------
# -------------------------
filename = 'MMO_XMLImport_070319.xlsx'
df = read_excel(filename)
set_option('precision', 0)

# -------------------------
# -----2. Data Cleanup-----
# -------------------------

# a. Find and remove empty data
df.replace(' ', NaN, inplace=True)
df.dropna(subset=['Full Name'], inplace=True)
df.dropna(subset=['Address'], inplace=True)

# b. Product Code Updates
df['Product Code'].replace('MMO ENROLLKIT', 'PEK', inplace=True)
df['Product Code'].replace('MMO MGNFR', 'MAG', inplace=True)
df['Product Code'].replace('MMO UMG', 'UMG', inplace=True)
df['Product Code'].replace(NaN, 'UMG', inplace=True)

# c. Product Desc Updates
df['Product Desc'].replace(
    'Optional Supplemental Benefit (OSB) Fulfillment Kit',
    '(OSB) Fulfillment Kit', inplace=True)
df['Product Desc'].replace(
    NaN, 'Understanding Medicare Guide', inplace=True)

# d. Order Type Updates
df['Order Type'].replace('SalesCallCenter', 'Call Center', inplace=True)
df['Order Type'].replace('End User', 'WEB', inplace=True)
df['Order Type'].replace('CustomerCare', 'Customer Service', inplace=True)
df['Order Type'].replace(NaN, 'BRE', inplace=True)

# e. Bill To Region Updates
df['Bill To Region'].replace('Region 1', '1', inplace=True)
df['Bill To Region'].replace('Region 2', '2', inplace=True)
df['Bill To Region'].replace(NaN, '2', inplace=True)

# f. Plan Year Updates
df['PlanYear'].replace(NaN, now.year, inplace=True)

# g. Remove duplicates
df['Full Name'] = df['Full Name'].str.title()
df['Address'] = df['Address'].str.title()
df['City'] = df['City'].str.title()
df.drop_duplicates(['Full Name', 'Address'], inplace=True)

# Reset the index and start at 1 instead of 0
df.reset_index(drop=True, inplace=True)
df.index += 1

df.sort_values(by='Product Code', inplace=True)

# -------------------------
# -----3. Output Data------
# -------------------------

writer = ExcelWriter('{0}_rev.xlsx'.format(filename))
df.to_excel(writer, index=False)
writer.save()
