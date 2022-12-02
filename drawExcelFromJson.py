# coding=utf-8

# @Classname     drawExcelFromJson
# @Date          2022/11/8 5:39 下午
# @Author        wangcanhui
# @Email         canhui@fintec.ai
# @Description   TODO

import pandas as pd

dj = pd.read_json('./data.json')
end_update_time = dj.get('data').get('end_update_time')
highlist = dj.get('data').get('highlist')
lowlist = dj.get('data').get('lowlist')
hcount = dj.get('data').get('hcount')
lcount = dj.get('data').get('lcount')

header = ['省', '市', '区/街道', '楼栋/村']

writer = pd.ExcelWriter('./疫情风险区{}.xlsx'.format(end_update_time))

sheet_name = '高风险区{}'.format(hcount)
pdh = pd.json_normalize(highlist).drop(columns=['type', 'area_name'])
pdh.to_excel(writer, sheet_name=sheet_name, header=header)

sheet_name = '低风险区{}'.format(lcount)
pdl = pd.json_normalize(lowlist).drop(columns=['type', 'area_name'])
pdl.to_excel(writer, sheet_name=sheet_name, header=header)

writer.save()
writer.close()
