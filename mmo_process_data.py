"""
Steps:
1. Import XML data - Done
2. Output XML File
3. Import data entry
4. Output Do Not Contact
5. Append data to xml list
5. Cleanup data
6. Output data file(s)
"""
from mmo_import_xml import XmlImport
from pandas import DataFrame

xml_dir = '//Xmf-server/duke/Inter Office Mail/MMO XML Orders/'

xml_df = XmlImport(xml_dir)
xml_df.parse_xml()
xml_df.xml_to_xlsx()

