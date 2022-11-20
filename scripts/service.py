"""
Файл содержит вспомогательные функции

"""

import pandas as pd
import pandas.io.formats.excel


def get_filtered_df(excel, dict_warehouses, name_column_whs, skiprows=0):
    """
    Возвращает dataframe из excel, отфильтрованный по складам
    с заменой наименований складов
    """
    full_df = excel.parse(skiprows=skiprows)
    filter_df = full_df[full_df[name_column_whs].isin(
        list(dict_warehouses.keys())
    )]
    return filter_df.replace({name_column_whs: dict_warehouses})


def save_to_excel(path, df):
    """
    Записывает данные в excel с автоматическим подбором ширины столбцов
    """
    sheet = 'Лист1'
    num_row_header = 0
    height_header = 50
    min_wdh_col = 15

    # Убираем формат заголовка таблицы по умолчанию
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name=sheet, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet]
    cell_form = workbook.add_format({
        'text_wrap': True, 'bold': True,
        'align': 'center', 'valign': 'vcenter'
    })
    worksheet.set_row(num_row_header, height_header, cell_format=cell_form)

    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), min_wdh_col)
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet].set_column(col_idx, col_idx, column_width)

    writer.close()
