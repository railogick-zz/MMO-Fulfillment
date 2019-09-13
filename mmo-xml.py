# Python 3.7.2
from os import path, getcwd, listdir, mkdir
from shutil import move
from xml.etree.ElementTree import parse
from datetime import datetime

from pandas import ExcelWriter, DataFrame


class XmlImport:
    def __init__(self):
        self.files = [p for p in listdir(getcwd()) if p.endswith(".xml")]
        self.total_orders = []
        self.order = []
        self.df = DataFrame()
        self.header_dict = {'ShipToName': 'Full Name', 'ShipToAddress1': 'Address', 'ShipToCity': 'City',
                            'ShipToState': 'State', 'ShipToZip': 'Zip', 'ShipToAddress3': 'Zone',
                            'ProductCode': 'Product Code', 'ProductName': 'Product Desc', 'PromiseDate': 'Drop Date',
                            'OrderType': 'Order Type', 'OrderDate': 'Order Date', 'UserEmail': 'Email',
                            'ShipToAddress4': 'Phone', 'BillToRegion': 'Bill To Region', 'PlanYear': 'PlanYear',
                            'PlanType': 'PlanType', 'MemberType': 'MemberType',
                            'WebtrendscampaignIDcode': 'WebtrendscampaignIDcode'}

    def parse_xml(self):
        for z in range(len(self.files)):
            tree = parse(self.files[z])
            root = tree.getroot()

            # Get information from each order and add to total_orders
            for each in root.iter('Order'):
                for x in range(len(self.header_dict)):
                    self.order.append(each.find(f'.//{list(self.header_dict.keys())[x]}').text)

                # Add completed order to total_orders and reset order list
                self.total_orders.append(self.order)
                self.order = []

        # Create Data frame from completed total_orders
        self.df = DataFrame(self.total_orders, columns=self.header_dict.values())

        now = datetime.now()
        folder = f'XML {now:%m%d%y}'
        if not path.exists(folder):
            mkdir(folder)
        for f in range(len(self.files)):
            move(self.files[f], folder)

        return self.df


def main():
    now = datetime.now()
    filename: str = f'MMO_XMLImport_{now:%m%d%y}.xlsx'
    import_xml = XmlImport()
    df = import_xml.parse_xml()
    writer = ExcelWriter(filename)
    df.to_excel(writer, index=False)
    writer.save()


if __name__ == '__main__':
    main()
