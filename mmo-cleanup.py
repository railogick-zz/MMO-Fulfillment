# Python 3.7.2
import re
from datetime import datetime

from numpy import NaN
from pandas import set_option, read_excel, ExcelWriter


class ProcessFile:
    def __init__(self, filename):
        set_option('precision', 0)
        self.file = filename
        self.df = read_excel(self.file)
        self.updates = {}

        # Find and remove empty data
        self.df.replace(' ', NaN, inplace=True)
        self.df.dropna(subset=['Full Name'], inplace=True)
        self.df.dropna(subset=['Address'], inplace=True)

    def update(self):
        # Dictionary of Updates
        self.updates = {'Product Code': {'MMO ENROLLKIT': 'PEK',
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
                        'PlanYear': {NaN: datetime.now().year}
                        }

        self.df.replace(self.updates, inplace=True)

    def remove_dupes(self):
        fix_cols = ['Full Name', 'Address', 'City']
        self.df[fix_cols] = self.df[fix_cols].applymap(lambda x: x.title())
        self.df[fix_cols] = self.df[fix_cols].applymap(lambda x: re.sub(' +', ' ', x))
        self.df.drop_duplicates(['Full Name', 'Address'], inplace=True)

        # Start the index at 1 and sort by 'Product Code'
        self.df.reset_index(drop=True, inplace=True)
        self.df.index += 1
        self.df.sort_values(by='Product Code', inplace=True)

    def return_df(self):
        self.update()
        self.remove_dupes()
        return self.df


def main():
    now = datetime.now()
    filename = f'MMO_XMLImport_{now:%m%d%y}.xlsx'
    job = ProcessFile(filename)
    df = job.return_df()
    writer = ExcelWriter(f'{filename[:-5]}_rev.xlsx')
    df.to_excel(writer, index=False)
    writer.save()


if __name__ == '__main__':
    main()
