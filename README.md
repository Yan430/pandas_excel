# pandas_excel
# How to generate well formatted (including grouping) Excel spreadsheet using Python and Pandas and xlsxwriter
  
This little snippet of code shows how you can get well formatted (including groping) Excel document from any table based data.

Original *.csv file - https://drive.google.com/file/d/1l2nCcx-Ew1bVr0gSYL3NMUjSQ8lRz_aw/view?usp=sharing

Excel document produced by the code - https://drive.google.com/file/d/1yQhqHna94XnKPAp7U2v-zVDf1N7diy4h/view?usp=sharing

Starting with provided *.csv file (part of Panda's Python package) - the code reads *.csv into Panda Dataframe (Banks_df), then creates second DataFrame (Grouped_ST_CT_df) that represents sorted by state groups of records (rows).

Iterating over groups in Grouped_ST_CT_df the code creates groups based on City and that is how records are added to resulting Dataframe (df_frame_out).

Excel grouping is done based on "LEVEL" column in df_frame_out - State names rows are assigned level "0", City names rows are assigned level "1", and rest of records are "2".

The code uses xlsxwriter Python package to generate final Excel document.

Function "set_row_groping_based_on_level" is used to generate Excel groping based on value of "LEVEL" column.

To execute this code simply run - python excel_gen1.py. Make sure that xlsxwriter is installed.
