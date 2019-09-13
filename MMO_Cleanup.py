# Python 3.7.2
import re
from datetime import datetime

from numpy import NaN
from pandas import set_option, read_excel, ExcelWriter

# TODO: Fix file import (GUI or search 'XMLImport')

now = datetime.now()
year, week_num, dow = now.isocalendar()  # Current year, week number, day of week

filename = 'MMO_XMLImport_082619.xlsx'
df = read_excel(filename)
set_option('precision', 0)

# Find and remove empty data
df.replace(' ', NaN, inplace=True)
df.dropna(subset=['Full Name'], inplace=True)
df.dropna(subset=['Address'], inplace=True)

# Dictionary of Updates
updates = {'Product Code': {'MMO ENROLLKIT': 'PEK',
                            'MMO MGNFR': 'MAG',
                            'MMO UMG': 'UMG',
                            NaN: 'UMG'},
           'Product Desc': {'Optional Supplemental Benefit (OSB) Fulfillment Kit': '(OSB) Fulfillment Kit',
                            NaN: 'Understanding Medicare Guide'},
           'Order Type': {'SalesCallCenter': 'Call Center',
                          'End User': 'WEB',
                          'CustomerCare': 'Customer Service',
                          NaN: 'BRE'},
           'Bill To Region': {'Region 1': '1',
                              'Region 2': '2',
                              NaN: '2'},
           'PlanYear': {NaN: now.year}
           }

df.replace(updates, inplace=True)
# for key, values in updates.items():
#     for val in values:
#         df[key].replace(val, values, inplace=True)

# Remove duplicates
fix_cols1 = ['Full Name', 'Address']
fix_cols2 = ['Full Name', 'Address', 'City']
df[fix_cols2] = df[fix_cols2].applymap(lambda x: x.title())
df[fix_cols1] = df[fix_cols1].applymap(lambda x: re.sub(' +', ' ', x))
df.drop_duplicates(fix_cols1, inplace=True)

# Reset the index and start at 1 instead of 0
df.reset_index(drop=True, inplace=True)
df.index += 1
df.sort_values(by='Product Code', inplace=True)

writer = ExcelWriter(f'{filename[:-5]}_rev.xlsx')
df.to_excel(writer, index=False)
writer.save()
