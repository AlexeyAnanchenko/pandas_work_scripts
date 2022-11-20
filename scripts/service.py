"""
Файл содержит вспомогательные функции

"""


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
