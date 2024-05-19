import pymongo
from openpyxl import Workbook
import pandas as pd

def getOrgInContent(content, orgInfo) -> list:
    orgList = list()
    for name, shorts in orgInfo.items():
        if name in content:
            orgList.append(name)
            continue
        for short in shorts:
            if short in content:
                orgList.append(name)
                break
    return orgList



def itemFilter(collection, orgInfo) -> dict:
    result = list()
    for item in collection.find():
        if item.get('data') is None:
            continue
        
        content = item['data']['content']
        org = getOrgInContent(content, orgInfo)
        if '省人民政府办公厅' in org and len(org) >= 3 or '省人民政府办公厅' not in org and len(org) >= 2:
            data = item['data'] 
            data['department'] = ','.join(org)
            data.pop('content')
            
            if data.get('pubDate') == None:
                data['pubDate'] = data['publish'].split(' ')[0]
            result.append(data)
            print(data)
    return result

def loadOrgInfo(path)->dict:
    result = dict()
    with open(path, 'r', encoding='utf-8') as file:
        for line in file.readlines():
            line = line.strip()
            name_short = line.split('：')

            key = name_short[0]
            values = []
            values.append(key)
            
            if '省' in key :
                values.append(key.replace('省', ''))
            if len(name_short) > 1:
                all_shorts = name_short[1].split('、')
                for short in all_shorts:
                    values.append(short)
                    if '省' in short :
                        values.append(short.replace('省',''))
            result[key] = values
            print(result)
    return result

def outputExcel(dictList, fileName):
    df = pd.DataFrame(dictList)
    df['date'] = pd.to_datetime(df['pubDate'])
    df['year'] = df['date'].dt.year
    excel_writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
    for year, group in df.groupby('year'):
        sheet_name = str(year)
        group.drop(columns=['date', 'year']).to_excel(excel_writer, sheet_name=sheet_name, index=False)
    excel_writer.close()

if __name__ == "__main__":
    client = pymongo.MongoClient("mongodb://localhost:27017/")
    db = client["GYDatabase"]
    xiamenGov = db['HebeiGovRegulations']
    
    orgInfo = loadOrgInfo('.deptList/Hebei.txt')
    outputExcel(itemFilter(xiamenGov, orgInfo),'./output/HebeiGovRegulations.xlsx')
    