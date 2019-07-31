# Python 3.7.2
from os import path, getcwd, listdir, mkdir
from shutil import move
from xml.etree.ElementTree import parse
from datetime import datetime

from pandas import ExcelWriter, DataFrame

# Establish global lists and variables

total_orders = []  # List of Orders to be converted to DataFrame
order = []  # Place holder list for individual orders.
now = datetime.now()  # Assign today to variable

# Dictionary of xml headers and corresponding header values
header_dict = dict(ShipToName='Full Name', ShipToAddress1='Address', ShipToCity='City', ShipToState='State',
                   ShipToZip='Zip', ShipToAddress3='Zone', ProductCode='Product Code', ProductName='Product Desc',
                   PromiseDate='Drop Date', OrderType='Order Type', OrderDate='Order Date', UserEmail='Email',
                   ShipToAddress4='Phone', BillToRegion='Bill To Region', PlanYear='PlanYear', PlanType='PlanType',
                   MemberType='MemberType', WebtrendscampaignIDcode='WebtrendscampaignIDcode')

# ----------------------
# -----1.Parse Data-----
# ----------------------

# Get list of xml files to parse.
files = [p for p in listdir(getcwd()) if p.endswith(".xml")]

# Parse through each file and extract orders
for z in range(len(files)):
    tree = parse(files[z])
    root = tree.getroot()

    # Get information from each order and add to total_orders
    for each in root.iter('Order'):
        for x in range(len(header_dict)):
            assert isinstance(each.find(f'.//{list(header_dict.keys())[x]}').text, object)
            order.append(each.find(f'.//{list(header_dict.keys())[x]}').text)

        # Add completed order to total_orders and reset order list
        total_orders.append(order)
        order = []

# Create Data frame from completed total_orders
df = DataFrame(total_orders, columns=header_dict.values())

# --------------------------
# ------2.Output Data-------
# --------------------------

# print(df[['Full Name','Address','Product Code']])
filename: str = f'MMO_XMLImport_{now:%m%d%y}.xlsx'
writer = ExcelWriter(filename)
df.to_excel(writer, index=False)
writer.save()

# --------------------------
# ------3.Move Files--------
# --------------------------

# Move completed files to dated folder
folder = f'XML {now:%m%d%y}'
if not path.exists(folder):
    mkdir(folder)
for f in range(len(files)):
    move(files[f], folder)