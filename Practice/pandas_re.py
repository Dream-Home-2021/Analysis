import pandas as pd

arr1 = pd.Series([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
# print(arr1)

arr2 = pd.DataFrame([[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20]])
# print(arr2)

arr3 = pd.DataFrame([[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20]], index=['a', 'b'])
# print(arr3)

arr4 = pd.DataFrame({'a': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 'b': [11, 12, 13, 14, 15, 16, 17, 18, 19, 20]})
# print(arr4)

df = pd.DataFrame({
    'int_col': [1, 2, 3],
    'float_col': [ '', "2.2", 3.3],
    'tr_col': ['a', 'b', 'c'],
    'datetime_col': ['', '2021-01-02', '2021-01-03']
})
# print(df.info())
# row = df.loc[0]
# print(row)

col = df.iloc[:2, :2]
# print(col)

t = df.loc[:, ['tr_col', 'datetime_col']]
# print(t)

a = df.loc[1:2, "float_col"]
print(df['float_col'])
print(df['float_col'].dtype)
df['float_col'] = df['float_col'].astype('float64')
print(df['float_col'].dtype)
print(df['float_col'])
# print(a)
# for i, j in df.iterrows():
#     print(i, j)
# 创建一个datetime64类型的时间序列
# dates = pd.date_range('20210101', periods=3)
#
# # 将datetime64类型转换为int64类型
# int_dates = dates.astype('int64')
#
# print(int_dates)