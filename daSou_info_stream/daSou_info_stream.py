# encoding=utf-8
import pandas as pd
from 大搜信息流.daSou_info_stream.Common.Common import StatisticsCalculator, data_1, data_2, data_3, stream, dasou, dasou_stream

# Load the data
df_main = pd.read_excel("Data/Book2.xlsx", sheet_name="Sheet1")
df_cs = pd.read_excel("Data/Book2.xlsx", sheet_name="客服通讯录(每月更新)")
df_order = pd.read_excel("Data/Book2.xlsx", sheet_name="订单行X")

# 数据清洗
df_main["部门"] = pd.Series(dtype="object")
df_main = pd.merge(df_main, df_cs[["管理员", "部门"]], on="管理员", how="left", suffixes=('_main', '_cs'))
df_main.drop(columns=["部门_main"], inplace=True)
df_main.rename(columns={"部门_cs": "部门"}, inplace=True)

# 大搜消费和信息流消费
columns = df_main.columns
date_columns = list(set([col[-8:] for col in columns if col[-8:].isdigit()]))
date_columns = sorted(date_columns)
dates = ["第1天", "第2天", "第3天", "第4天", "第5天", "第6天", "第7天", "第8天"]
for date, i in zip(date_columns, dates):
    df_main[f"大搜消费{i}"] = df_main[f"总消费{date}"] - df_main[f"原生自主投放总消费{date}"] - df_main[
        f"凤巢优惠券消费{date}"] - df_main[f"聚屏平台合约消费{date}"] - df_main[f"度星选-软植互选-消费{date}"]

    df_main[f"信息流消费{i}"] = df_main[f"原生自主投放总消费{date}"] - df_main[f"原生CPC优惠券消费{date}"] - df_main[
        f"原生CPM优惠券消费{date}"]

# "城市&框架"
df_main["城市&框架"] = df_main["部门"].apply(lambda x: "框架" if x == "框架" else None)

df_main["城市&框架"].where(df_main["城市&框架"].notna(), df_main["发证机关所在市"].apply(
    lambda x: x if x in ["厦门市", "泉州市", "漳州市", "龙岩市"] else None), inplace=True)

df_main["城市&框架"].where(df_main["城市&框架"].notna(), df_main["订单行"].apply(
    lambda x: x + "市" if x in ["厦门", "泉州", "漳州", "龙岩"] else "其他市"), inplace=True)

# Calculate "大搜7日均", "信息流7日均" and "大搜+信息流消费近7天"
columns = df_main.columns
daSou_columns = [col for col in columns if '大搜消费' in col][-7:]
infoStream_columns = [col for col in columns if '信息流消费' in col][-7:]
df_main["大搜7日均"] = df_main[daSou_columns].mean(axis=1)
df_main["信息流7日均"] = df_main[infoStream_columns].mean(axis=1)
df_main["大搜+信息流消费近7天"] = df_main[daSou_columns].sum(axis=1) + df_main[infoStream_columns].sum(axis=1)


calculator = StatisticsCalculator(df_main)

# 城市和部门维度--大搜消费和信息流统计
department_stats, city_stats = calculator.calculate()

# 更新年框缓存数据
calculator.update_cache(data_1, data_2, data_3)

# 缓存首次导入
# calculator.save_data_to_cache(dasou, 'cache_0.pkl')
# calculator.save_data_to_cache(dasou_stream, 'cache_2.pkl')
# calculator.save_data_to_cache(stream, 'cache_1.pkl')


# df1 = pd.DataFrame(department_stats).T.round(0)
# df2 = pd.DataFrame(city_stats).T.round(0)
#
# with pd.ExcelWriter('output.xlsx') as writer:
#     df1.to_excel(writer, sheet_name='Sheet1')
#     df2.to_excel(writer, sheet_name='Sheet2')
#
# print(result)
