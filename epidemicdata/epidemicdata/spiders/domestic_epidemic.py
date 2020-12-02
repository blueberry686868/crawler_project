# -*- coding: utf-8 -*-
import scrapy
import requests
import json   # 处理数据
from openpyxl import load_workbook    # 保存为表格


class DomesticEpidemicSpider(scrapy.Spider):
    name = 'domestic_epidemic'
    allowed_domains = ['view.inews.qq.com']
    start_urls = ['https://view.inews.qq.com/g2/getOnsInfo?name=disease_h5']

    def parse(self, response):
        url = self.start_urls[0]

        yield scrapy.Request(
            url=url,
            callback=self.parse_one  # 响应结果回调到parse_one函数中
        )

    def parse_one(self, response):

        # 得到国内的总体数据
        result = json.loads(json.loads(response.text)['data'])['areaTree']

        # 得到国内各省市的子数据（这里以云南为例）
        data = result[0]['children']
        for i in data:
            if i['name'] == '云南':
                data1 = i

                # 获取市区名字、今日新增人数、累计确诊人数
                name_yunnan = []
                total_yunnan = []
                new_yunnan = []
                for j in data1['children']:
                    name_yunnan = name_yunnan + [j['name']]
                    total_yunnan = total_yunnan + [j['total']]
                    new_yunnan = new_yunnan + [j['today']]

                today_yn = []
                for item in new_yunnan:
                    today_yn = today_yn + [item['confirm']]

                total_yn = []
                for item in total_yunnan:
                    total_yn = total_yn + [item['confirm']]

                print(data1['children'])  # 云南省的各市区信息
                print(len(data1['children']))  # 云南省有几条信息
                print(name_yunnan)
                print(today_yn)
                print(total_yn)

                workbook = load_workbook(filename=r"/home/ubuntu/Desktop/疫情数据.xlsx")

                sheet = workbook.active

                sheet["A1"] = "市区名字"
                sheet["B1"] = "今日新增人数"
                sheet["C1"] = "累计确诊人数"
                sheet["E1"] = "市区名字"
                sheet["F1"] = "今日新增人数"
                sheet["G1"] = "累计确诊人数"

                # 保存到表格的A、B、C列
                # 云南
                j = 2
                while j <= 16:
                    sheet["A%d" % j] = name_yunnan[j - 2]
                    sheet["B%d" % j] = today_yn[j - 2]
                    sheet["C%d" % j] = total_yn[j - 2]
                    j += 1

                # 北京
                # k = 2
                # while k <= 19:
                #     sheet["E%d" % k] = name_beijing[k - 2]
                #     sheet["F%d" % k] = today_bj[k - 2]
                #     sheet["G%d" % k] = total_bj[k - 2]
                #     k += 1

                workbook.save(filename=r"/home/ubuntu/Desktop/疫情数据.xlsx")
                print("保存成功")




