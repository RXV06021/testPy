# -*- coding: utf-8 -*-
"""
Created on Sat Apr 21 12:03:11 2018
@author: 孝博

csvデータを入手し所定のエクセルsheetにコピーする
"""

# import lib
import pandas as pd
import urllib.request

# import csv data
url = "https://indexes.nikkei.co.jp/nkave/historical/nikkei_stock_average_monthly_jp.csv"  # ex.日経平均月次データ
savename = "testData.csv"                                                                  # ダウンロードファイルの名称
urllib.request.urlretrieve(url, savename)                                                  # ファイルのダウンロード実行
df = pd.read_csv("testData.csv", index_col=0, encoding="shift_jis", skipfooter=1, engine='python')

#%%

df.to_excel("testData1.xlsx",startrow=0,startcol = 0)

#%%
# Create a Pandas Excel writer using XlsxWriter as the engine.
# Convert the dataframe to an XlsxWriter Excel object.
writer = pd.ExcelWriter('testData2.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='SheetA')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#%%
# Create a Pandas Excel writer using XlsxWriter as the engine.
# Convert the dataframe to an XlsxWriter Excel object.
writer = pd.ExcelWriter("testData3.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name='SheetB')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['SheetB']

# Add some cell formats.
fmt1 = workbook.add_format()
fmt2 = workbook.add_format({'num_format': '#,##0.0',
                            'border': 0})
      # Note: It isn't possible to format any cells that already have a format such
      #       as the index or headers or any cells that contain dates or datetimes.
# Set the column width and format.
worksheet.set_column('A:A', 11, fmt1)
worksheet.set_column('B:E', 10, fmt2)

# Add a header format.
header_fmt = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#CCFFCC',
    'border': 1})
# Write the column headers with the defined format.
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num + 1, value, header_fmt)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
