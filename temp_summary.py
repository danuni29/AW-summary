import openpyxl as op
from copy import copy
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows


def main():
    df = pd.read_csv('AW_summary_all_20200811.csv')
    wb = op.load_workbook('summary_template_20200812_5개년.xlsx')
    ws = wb['용수구역(통합표)']

    rows = ws.iter_rows(min_row=5, max_row=5)
    for row in rows:
        for cell in row:
            new_cell = ws.cell(row=7, column=cell.col_idx, value=cell.value)

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)



    df1 = df[df['용수구역명'] == '강릉1']

    locations = list(df['용수구역명'].unique())
    print(len(locations))

    five_years_temp = []
    for location in locations:
        cut_df = df[df['용수구역명'] == location]
        five_years_df = cut_df[(cut_df['연도'] >= 2011) & (cut_df['연도'] <= 2015)]
        common_years_df = cut_df[(cut_df['연도'] >= 1981) & (cut_df['연도'] <= 2010)]
        # print(five_years_df)
        five_years_temp.append(sum(five_years_df['연평균 일평균기온(℃)']) / len(five_years_df['연평균 일평균기온(℃)']))


        for r in dataframe_to_rows(five_years_df, index=False, header=False):
            print(five_years_df['연평균 일평균기온(℃)'])
            ws.append(five_years_temp)

    print(five_years_temp)







    wb.save('summary_template_20200812_5개년.xlsx')


if __name__ == '__main__':
    main()