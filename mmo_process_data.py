from mmo_import_xml import XmlImport
from mmo_fix_data import ProcessFile
from gooey import Gooey, GooeyParser
from pandas import ExcelWriter, read_excel, concat
from datetime import datetime

now = datetime.now()


def output_contact_dnc(contact_df, dnc_df):
    filename = f'MMO CONTACT & DO NOT MAIL {now:%m-%d-%Y}.xlsx'
    writer = ExcelWriter(filename, engine='xlsxwriter')
    contact_df.to_excel(writer, sheet_name='Sheet1', startrow=2, header=False, index=False)
    wb = writer.book
    ws = writer.sheets['Sheet1']
    ws.set_column('A:H', 20)
    ws.write('A1', f'{now:%m/%d/%Y}')
    contact_df.columns = map(str.upper, contact_df.columns)
    ws.write_row('A2', list(contact_df.columns.values))
    if not dnc_df.empty:
        dnc_row = len(contact_df.index) + 4
        ws.write(dnc_row, 0, 'DO NOT MAIL')
        dnc_df.to_excel(writer, sheet_name='Sheet1', startrow=dnc_row + 1, header=False, index=False)

    writer.save()


@Gooey
def main():
    parser = GooeyParser()
    parser.add_argument('-Data_File',
                        help="Select data entry file",
                        widget="FileChooser")
    args = parser.parse_args()

    # Process XML file
    xml_dir = '//Xmf-server/duke/Inter Office Mail/MMO XML Orders/'
    mmo_xml = XmlImport(xml_dir)
    mmo_xml.parse_xml()
    mmo_xml.xml_to_xlsx()

    # Create main working DataFrame
    mmo_df = mmo_xml.df

    # Process Data Entry file if present
    if args.Data_File:
        data_entry_df = read_excel(args.Data_File, dtype=str)
        df_contact = data_entry_df.loc[data_entry_df['STATUS'] == 'CONTACT'].drop(columns=['STATUS']).reset_index(drop=True)
        df_dnc = data_entry_df.loc[data_entry_df['STATUS'] == 'DNC'].drop(columns=['STATUS']).reset_index(drop=True)
        df_data = data_entry_df.loc[data_entry_df['STATUS'] == 'Data'].drop(columns=['STATUS']).reset_index(drop=True)
        #
        output_contact_dnc(df_contact, df_dnc)
        mmo_df = concat([mmo_df, df_contact, df_data],
                        ignore_index=True, sort=False).drop(columns=['Check Box']).fillna('')

    job = ProcessFile(mmo_df)


if __name__ == '__main__':
    main()
