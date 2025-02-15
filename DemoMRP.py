# 使用封包
import pandas as pd
import numpy as np
import os
import openpyxl
import time

start = time.time()
print("程式開始執行...", "\n")

pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.expand_frame_repr", None)

path = os.getcwd()
df = pd.read_excel(r"C:\Users\admin\Desktop\MRP.XLSX", engine="openpyxl")

print("Excel資料成功轉換成dataframe...", "\n")

col_list = [i for i in df.columns]
condition = df["REQUEST ITEM"] == "FIRM ORDERS"

dff = df.loc[condition, ["PARTNO", "REQUEST ITEM", "TOTAL"]]
dff.columns = ["PARTNO", "REQUEST ITEM", "TOTAL_FIRM ORDERS"]
dff["REQUEST ITEM"] = "GROSS REQTS"

dff.index = [i for i in range(len(dff))]
dfc = pd.merge(df, dff, on=["PARTNO", "REQUEST ITEM"], how="left")

condition = df["REQUEST ITEM"] == "GROSS REQTS"
dfc.loc[condition, "LEAD_TIME_Week"] = np.ceil(dfc["LEAD TIME"] / 7)
start_col = dfc.columns.get_loc("PASSDUE")
end_col = dfc.columns.get_loc("FUTURE")

# 計算MRP報表長度(PASSDUE到FUTURE間，所跨越的週數)
term_week_col = end_col - start_col + 1
for i in range(len(dfc)):
    if dfc.loc[i, "LEAD_TIME_Week"] >= term_week_col:
        dfc.loc[i, "LEAD_TIME_Week_Limit"] = term_week_col
    if dfc.loc[i, "LEAD_TIME_Week"] < term_week_col:
        dfc.loc[i, "LEAD_TIME_Week_Limit"] = dfc.loc[i, "LEAD_TIME_Week"]


dfc.loc[condition, "Total_Stock"] = (
    dfc.loc[condition, "PLANT STOCK"] + dfc.loc[condition, "INSP. STOCK"]
)

selected_columns = ["PARTNO", "REQUEST ITEM", "Total_Stock"] + list(
    dfc.loc[:, "PASSDUE":"FUTURE"].columns
)
dfc_result = dfc.loc[condition, selected_columns]
dfc_result.reset_index(drop=True, inplace=True)
gross_cum_list = list(dfc.loc[:, "PASSDUE":"FUTURE"].columns)


for i in range(len(dfc_result)):
    res = 0
    for j in gross_cum_list:
        res = res + dfc_result.loc[i, j]
        if res > dfc_result.loc[i, "Total_Stock"]:
            dfc_result.loc[i, "Demand_Week"] = j
            break
        else:
            dfc_result.loc[i, "Demand_Week"] = "no demand"

dfc = pd.merge(
    dfc,
    dfc_result.reindex(columns=["PARTNO", "REQUEST ITEM", "Demand_Week"]),
    on=["PARTNO", "REQUEST ITEM"],
    how="left",
)

selected_columns = ["PARTNO", "REQUEST ITEM", "LEAD_TIME_Week_Limit"] + list(
    dfc.loc[:, "PASSDUE":"FUTURE"].columns
)

dfcc = dfc.loc[condition, selected_columns]
dfcc.reset_index(drop=True, inplace=True)
start_col = dfcc.columns.get_loc("PASSDUE")

for i in range(len(dfcc)):
    res = 0
    for j in range(int(dfcc.loc[i, "LEAD_TIME_Week_Limit"])):
        res = res + dfcc.iloc[i, start_col + j]
    dfcc.loc[i, "LEAD_TIME_GrossRequest"] = res


dfc = pd.merge(
    dfc,
    dfcc[["PARTNO", "REQUEST ITEM", "LEAD_TIME_GrossRequest"]],
    on=["PARTNO", "REQUEST ITEM"],
    how="left",
)


dfc.loc[condition, "AMT"] = (
    dfc.loc[condition, "Total_Stock"]
    + dfc.loc[condition, "TOTAL_FIRM ORDERS"]
    - dfc.loc[condition, "LEAD_TIME_GrossRequest"]
)
dfc["PARTNO"] = dfc["PARTNO"].astype(str)

select_columns = (
    ["PURCHASING GROUP", "REQUEST ITEM"]
    + list(dfc.loc[:, "PASSDUE":"FUTURE"].columns)
    + ["TOTAL_FIRM ORDERS", "Total_Stock"]
)
dfg_1 = dfc.groupby(select_columns[:2])[select_columns[2:]].sum()
dfg_1.reset_index(inplace=True)

select_columns = (
    ["PURCHASING GROUP", "REQUEST ITEM"]
    + list(dfc.loc[:, "PASSDUE":"FUTURE"].columns)
    + ["LEAD_TIME_Week", "LEAD_TIME_Week_Limit"]
)
dfg_2 = dfc.groupby(select_columns[:2])[select_columns[2:]].min()
dfg_2.reset_index(inplace=True)

dfg_c = pd.merge(
    dfg_1,
    dfg_2[
        ["PURCHASING GROUP", "REQUEST ITEM", "LEAD_TIME_Week", "LEAD_TIME_Week_Limit"]
    ],
    on=["PURCHASING GROUP", "REQUEST ITEM"],
    how="left",
)


for i in range(len(dfg_c)):
    res = 0
    for j in gross_cum_list:
        res = res + dfg_c.loc[i, j]
        if res > dfg_c.loc[i, "Total_Stock"]:
            dfg_c.loc[i, "Demand_Week"] = j
            break
        else:
            dfg_c.loc[i, "Demand_Week"] = "no demand"
dfg_c.loc[
    dfg_c["REQUEST ITEM"] == "FIRM ORDERS", "LEAD_TIME_Week_Limit":"Demand_Week"
] = np.nan
dfg_c.loc[
    dfg_c["REQUEST ITEM"] == "NET  AVAIL", "LEAD_TIME_Week_Limit":"Demand_Week"
] = np.nan
dfg_c.loc[
    dfg_c["REQUEST ITEM"] == "PLAN ORDERS", "LEAD_TIME_Week_Limit":"Demand_Week"
] = np.nan


for i in range(len(dfg_c)):
    res = 0
    try:
        for j in range(int(dfg_c.loc[i, "LEAD_TIME_Week_Limit"])):
            res = res + dfg_c.iloc[i, 2 + j]

        dfg_c.loc[i, "LEAD_TIME_GrossRequest"] = res
    except:
        continue

condition = dfg_c["REQUEST ITEM"] == "GROSS REQTS"
dfg_c.loc[condition, "AMT"] = (
    dfg_c.loc[condition, "Total_Stock"]
    + dfg_c.loc[condition, "TOTAL_FIRM ORDERS"]
    - dfg_c.loc[condition, "LEAD_TIME_GrossRequest"]
)

dfc = pd.concat([dfc, dfg_c], axis=0)

dfc.to_csv(path + "\\" + "result.txt", index=False, encoding="utf-8-sig", sep="\t")

end = time.time()
res = end - start
print(f"程式執行結束...耗時{res:.2f}秒", "\n")
