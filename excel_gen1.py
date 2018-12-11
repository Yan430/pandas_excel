# example on how to generate well structured Excel file from any kind of row table form date

import os
import pandas as pd  # this is how I usually import pandas
from xlsxwriter.utility import xl_rowcol_to_cell


def format_excel(writer, df_size, dt_frame):


    workbook = writer.book
    worksheet = writer.sheets['banks_all']
    worksheet.outline_settings(True, False, True, False)

    format1 = workbook.add_format({'text_wrap': False})
    format1.set_align('left')
    format1.set_align('top')
    format1.set_font_name('Calibri')
    format1.set_font_size(11)


    worksheet.set_column('A:A', 5, format1)  # State
    worksheet.set_column('B:B', 20, format1)  # City
    worksheet.set_column('C:C', 45, format1)  # Bank Name
    worksheet.set_column('D:D', 10, format1)  # Certificate
    worksheet.set_column('E:E', 35, format1) # Acquiring Institution
    worksheet.set_column('F:F', 15, format1) # Closing Date
    worksheet.set_column('G:G', 15, format1)  # Updated Date
    worksheet.set_column('H:H', 3, format1)  # level


    dt_frame.apply(set_row_groping_based_on_level, axis=1, args=(worksheet,format1))


    # Add 1 to row so we can include a total
    # subtract 1 from the column to handle because we don't care about index
    table_end = xl_rowcol_to_cell(df_size[0], df_size[1] - 1)
    table_range = 'A1:{}'.format(table_end)

    """
        Bank Name	City	ST	CERT	Acquiring Institution	Closing Date	Updated Date Level 
        
    """


    worksheet.add_table(table_range, {'columns': [{'header': 'State'},
                                                  {'header': 'City'},
                                                  {'header': 'Bank Name'},
                                                  {'header': 'CERT'},
                                                  {'header': 'Acquiring Institution'},
                                                  {'header': 'Closing Date'},
                                                  {'header': 'Updated Date'},
                                                   {'header': 'Level'}],
                                      'style': 'Table Style Light 9'})

def set_row_groping_based_on_level(row, worksheet, cell_format_in):
    row_index = int(row.name) + 1
    row_level = int(row['LEVEL'])
    cell_format = cell_format_in

    if row_level == 0:
        worksheet.set_row(row_index,  None, cell_format, {'level': row_level, 'collapsed': True})
    else:
        worksheet.set_row(row_index, None, cell_format, {'level': row_level,'hidden': True})



def main():

    pd.set_option('display.max_colwidth', -1)

    cwd = os.getcwd()
    csv_path = cwd + '\\' + 'banklist.csv'
    Banks_df = pd.read_csv(csv_path)
    df_frame_out = pd.DataFrame(columns=Banks_df.columns, dtype=Banks_df.dtypes)

    df_frame_out['LEVEL'] = '2'


    FileName_out = cwd + '\\' + 'banklist_rep.xlsx'

    Grouped_ST_CT_df = Banks_df.groupby(['ST'])

    for name, group in Grouped_ST_CT_df:
        new_row_df = pd.DataFrame([[''] * 8], columns=df_frame_out.columns)
        new_row_df['ST'] = name
        new_row_df['LEVEL'] = '0'

        df_frame_out = df_frame_out.append(new_row_df)

        Grouped_CT_df = group.groupby(['City'])

        for ct_name, ct_group in Grouped_CT_df:
            new_row_df1 = pd.DataFrame([[''] * 8], columns=df_frame_out.columns)
            new_row_df1['City'] = ct_name
            new_row_df1['LEVEL'] = '1'
            df_frame_out = df_frame_out.append(new_row_df1)

            for index, row in ct_group.iterrows():
                row['LEVEL'] = '2'
                df_frame_out = df_frame_out.append(row)

    df_frame_out = df_frame_out.reset_index(drop=True)

    df_frame_out = df_frame_out[['ST', 'City', 'Bank Name', 'CERT', 'Acquiring Institution', 'Closing Date', 'Updated Date', 'LEVEL']]

    writer = pd.ExcelWriter(FileName_out, engine='xlsxwriter')

    df_frame_out.to_excel(writer, 'banks_all', index=False)

    format_excel(writer, df_frame_out.shape, df_frame_out)

    writer.save()

    print 'completed - please check ' + FileName_out


if __name__ == '__main__':
    main()

