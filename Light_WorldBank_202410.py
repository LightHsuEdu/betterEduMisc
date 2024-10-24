# -*- coding: utf-8 -*-

# -----------------------------------------------------------------
#
# 使用说明：
# 1. 浏览器打开 https://data.worldbank.org.cn/indicator
# 2. 保存页面，注意保存到此 .py 文件所在文件夹
# 3. 运行此.py 文件
#
# -----------------------------------------------------------------

import os
import glob
from datetime import datetime
from urllib.request import urlopen
from bs4 import BeautifulSoup
import pandas as pd

# 获取当前 年月日时
current_datetime_string = datetime.now().strftime("%Y%m%d%H")

# 最终输出 excel 文件
outputFile = "WorldBank_Data_" + current_datetime_string +".xlsx"

# 以当前日期 年月日时 建立文件夹
wb_xls_subdirectory = current_datetime_string
try:
    os.mkdir(wb_xls_subdirectory)
    print(f"\n文件夹 '{wb_xls_subdirectory}' 成功建立.\n")
except FileExistsError:
    print(f"\n文件夹 '{wb_xls_subdirectory}' 已经存在.\n")
except Exception as err:
    print(f"错误: {err}")

# 读取当前文件夹下所有 .htm* 文件
htm_files = glob.glob("*.htm*")

# 数据文件 excel 格式下载链接
downloadlist = []

# 爬取 xls 文件保存目录
crawler_saveLocation = wb_xls_subdirectory + "/"

# 爬取并保存      
def saveWB_DataFile(linkurl):
    save_as = crawler_saveLocation + linkurl.replace("https://api.worldbank.org/v2/zh/indicator/", "").replace("?downloadformat=excel", "") + ".xls"
    with urlopen(linkurl) as file:
        content = file.read()
        with open(save_as, 'wb') as download:
            print("保存中...  " + linkurl)
            download.write(content)
      
# 提取保存页面的所有数据文件 excel 格式下载链接
for fileName in htm_files:
    try:
        with open(fileName,'r',encoding= 'utf-8') as flist:
            fileText = flist.read()
            indicatorsBlock = BeautifulSoup(fileText, "lxml")
            for tag in indicatorsBlock.find_all(href=True):
                linkurl = tag['href']
                if linkurl[:11] == "/indicator/":
                    full_url = "https://data.worldbank.org.cn" + linkurl
                    download_xlsUrl = "https://api.worldbank.org/v2/zh" + linkurl.replace("?view=chart", "?downloadformat=excel")                    
                    downloadlist.append(download_xlsUrl)
    except IOError as exc:
        if exc.errno != errno.EISDIR:
            raise

# 下载保存
for i in range(len(downloadlist)):
    saveWB_DataFile(downloadlist[i])

print("\n下载完成！\n开始转换数据......\n")

# 读取 xls 保存目录所有文件
wbXlsPath = wb_xls_subdirectory + '/*.xls'
wbFileLists = glob.glob(wbXlsPath)

# 用于输出 xls 的 DataFrame 
readyToXlsDF = pd.DataFrame()
firstSheetFile = True
fileNumCnt = 1

# 生成最终输出 xls 的 DataFrame 
for wbXlsFile in wbFileLists:
    sheet_1_df = pd.read_excel(wbXlsFile, sheet_name='Data', skiprows=3)
    # 取指标名
    var_name = sheet_1_df.iat[1, 2]
    # 删除不需要的列
    sheet_1_df.drop(inplace=True, columns=['Indicator Code','Country Name','Indicator Name'])
    # 取Income_Group列
    sheet_2_df = pd.read_excel(wbXlsFile, sheet_name='Metadata - Countries', usecols=[0,1,3])
    allSheets = pd.merge(sheet_2_df, sheet_1_df, on="Country Code")
    print('正在处理： ' + wbXlsFile)
    print('文件数量： ' + str(fileNumCnt))
    print('')
    fileNumCnt += 1
    
    if firstSheetFile == True:
        readyToXlsDF = pd.melt(allSheets, id_vars=['Country Name','Country Code','Income_Group'], var_name="Year", value_name='var_' + var_name)
        firstSheetFile = False
    else:
        allSheetReshape = pd.melt(allSheets, id_vars=['Country Name','Country Code','Income_Group'], var_name="Year", value_name='var_' + var_name)
        readyToXlsDF = pd.concat([readyToXlsDF, allSheetReshape], axis=1)
        readyToXlsDF = readyToXlsDF.loc[:,~readyToXlsDF.columns.duplicated()]

# 更改列名
readyToXlsDF.rename(columns={'Country Name':'Country_Name'}, inplace=True)

print('\n正在写入数据文件......')
readyToXlsDF.to_excel(outputFile, index=False)
print('\n转换完成！\n')

