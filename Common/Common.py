# coding: utf-8

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import numpy as np
import datetime
import openpyxl
import re


# filter类
class Filter:
    def __init__(self, df):
        self.data = df

    def filter_data(self, include_values=None, exclude_values=None, na_action='exclude',
                    date_column=None, date_range=None,
                    num_column=None, num_range=None):
        """
        :param include_values: dict, where key is the column name and value is a list of values to be filtered.
                           If not provided, defaults to filtering all values.
        :param exclude_values: dict, where key is the column name and value is a list of values to be excluded.
                               If NA is provided in values list, will exclude all NaN/None rows for that column.
        :param na_action: str, 'include' or 'exclude', default 'include'. If 'include', will include all NaN/None rows.
        ：param date_column: str, the column name of date column
        :param date_range: tuple, the start date and end date of date range
        :param num_range: str, like ">0", "<=100", "0-100" to indicate the range of filtering.
        :return: pd.DataFrame after filtering
        :return: pd.DataFrame after filtering
        :example:   filter_instance = Filter(df)
                    result1 = filter_instance.filter_data(include_values={'A': ['1', '2'], 'B': ["3"]}, na_action='include')
                    result2 = filter_instance.filter_data(exclude_values={'C': ['4']})
                    result1 = filter_instance.filter_data(include_values={'A': ['1', '2'], 'B': ["3"]}, exclude_values={'C': ['4']}, na_action='include')
                    result1 = filter_instance.filter_data(date_column='date', date_range=('2020-01-01', '2020-02-01'))
                    result1 = filter_instance.filter_data(date_column='date', date_range=('2020-01-01', None))
                    result1 = filter_instance.filter_data(num_column='num', num_range="1-100")
                    result1 = filter_instance.filter_data(num_column='num', num_range=">=100")

        """
        if include_values is None and exclude_values is None and date_range is None and num_range is None:
            return self.data

        result = self.data.copy()
        # 日期筛选功能
        if date_column and date_range:
            if date_column not in result.columns:
                raise ValueError(f"Date column '{date_column}' not found in dataframe.")

            start_date, end_date = date_range
            if start_date:
                result[date_column] = pd.to_datetime(result[date_column], errors='coerce')
                result = result[result[date_column] >= pd.to_datetime(start_date)]
            if end_date:
                result[date_column] = pd.to_datetime(result[date_column], errors='coerce')
                result = result[result[date_column] <= pd.to_datetime(end_date)]

        # 数值筛选功能
        if num_column and num_range:
            if num_column not in result.columns:
                raise ValueError(f"Numeric column '{num_column}' not found in dataframe.")

            # 处理数值列的数据
            result[num_column] = result[num_column].astype(str).str.replace(',', '')
            result[num_column] = result[num_column].str.extract('([-+]?\d*\.\d+|\d+)').astype(float)

            # 解析num_range字符串
            if ">" in num_range:
                if "=" in num_range:
                    limit = float(re.search(">=([\d.]+)", num_range).group(1))
                    result = result[result[num_column] >= limit]
                else:
                    limit = float(re.search(">([\d.]+)", num_range).group(1))
                    result = result[result[num_column] > limit]

            elif "<" in num_range:
                if "=" in num_range:
                    limit = float(re.search("<=([\d.]+)", num_range).group(1))
                    result = result[result[num_column] <= limit]
                else:
                    limit = float(re.search("<([\d.]+)", num_range).group(1))
                    result = result[result[num_column] < limit]

            elif "-" in num_range:
                lower_limit, upper_limit = map(float, num_range.split('-'))
                result = result[(result[num_column] >= lower_limit) & (result[num_column] <= upper_limit)]

        if include_values:
            for col, values in include_values.items():
                # 如果col不在result.columns中，抛出异常
                if col not in result.columns:
                    raise ValueError(f"Column '{col}' not found in dataframe.")
                # 遍历values中的每一个值，如果不是NA，就判断是否在result[col]中，如果不在，抛出异常. all()函数，如果有一个不在，就返回False
                if not all(v in result[col].values for v in values if v != "NA"):
                    raise ValueError(f"One or more values from {values} not found in column '{col}'.")
                # 如果NA在values中，就判断result[col]中是否有NaN，如果有，就根据na_action的值，决定是否筛选空值还是不筛选空值
                if "NA" in values:
                    if na_action == 'exclude':
                        result = result[result[col].notna()]
                        break
                    elif na_action == 'include':
                        result = result[result[col].isna()]
                        break

                result = result[result[col].isin(values)]

        if exclude_values:
            for col, values in exclude_values.items():
                if col not in result.columns:
                    raise ValueError(f"Column '{col}' not found in dataframe.")
                if not all(v in result[col].values for v in values if v != "NA"):
                    raise ValueError(f"One or more values from {values} not found in column '{col}'.")
                if "NA" in values:
                    """ result[col]：这部分表示选取 result 数据框中的特定列 col，得到一个表示该列的 Pandas Series。
                        .notna()：这是一个 Pandas Series 的方法，用于返回一个布尔值 Series，其中对于原始 Series 中的每个元素，如果元素不是 NaN（缺失值），则返回 True，否则返回 False。
                        result[col].notna()：通过将 .notna() 方法应用于列 col，你得到了一个布尔值 Series，其中的每个元素表示相应的行是否不是 NaN。
                        result[result[col].notna()]：通过将布尔值 Series 作为索引，你筛选出了 result 数据框中所有在列 col 中不是 NaN 的行，得到一个新的数据框。"""
                    result = result[result[col].notna()]
                    break
                result = result[~result[col].isin(values)]

        return result.reset_index(drop=True)


class Matcher:

    def __init__(self):
        pass

    def match(self, df1, df2, **kwargs):
        if not kwargs:
            raise ValueError("Please provide dependency and result columns.")

        dependency_col = kwargs.get("dependency", None)

        # 支持单个字符串或列表输入
        if isinstance(kwargs.get("result"), list):
            result_cols = kwargs.get("result")
        else:
            result_cols = [kwargs.get("result")]

        if not dependency_col:
            raise ValueError("Dependency column not provided.")
        if not result_cols:
            raise ValueError("Result column(s) not provided.")

        # 验证依赖列是否存在于数据帧中
        if dependency_col not in df1.columns:
            raise ValueError(f"Column '{dependency_col}' not found in first dataframe.")
        if dependency_col not in df2.columns:
            raise ValueError(f"Column '{dependency_col}' not found in second dataframe.")

        # 验证每一个结果列是否存在于df2中
        for col in result_cols:
            if col not in df2.columns:
                raise ValueError(f"Column '{col}' not found in second dataframe.")

        # 使用merge进行匹配
        cols_to_merge = [dependency_col] + result_cols
        result_df = pd.merge(df1, df2[cols_to_merge], on=dependency_col, how='left')

        # 对于每一个结果列，如果有冲突则解决命名冲突
        for col in result_cols:
            if col + '_x' in result_df.columns:
                result_df.drop(columns=[col + '_x'], inplace=True)
                result_df.rename(columns={col + '_y': col}, inplace=True)

        return result_df


if __name__ == '__main__':
    # pd.set_option('display.max_columns', None)
    # pd.set_option('display.max_rows', None)
    #
    # df = pd.read_excel('../通讯录.xlsx', usecols=[0, 1, 2, 3])
    # filter_instance = Filter(df)
    #
    # # Example 1: Filtering rows where column 'A' is either 'a' or 'b' and column 'B' is either 1 or 2
    # result1 = filter_instance.filter_data(conditions={'部门': ['大客部门', '框架'], '组别': ["行业二部"]})
    # print(result1)
    pass
