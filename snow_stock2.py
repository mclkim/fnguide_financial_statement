# -*- coding: utf-8 -*-
import math
from os import path
from os.path import join
from time import sleep

import requests
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
import re
from openpyxl import Workbook

from lxml import html

# fnguide 로 크롤링 변경예정..
# http://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701
#
# http://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701

# 눈덩이 주식계산을통한 추천종목 넣을 데이터프레임 생성
# sheet_row = ['회사명', 'BPS', '평균ROE', '현재가', '5년후 미래가치', '기대수익률', '투자가능가격', '점수']


# 종목코드 저장한 엑셀파일 불러오기(엑셀출처:[한국거래소 전자공시 홈페이지](http://kind.krx.co.kr/corpgeneral/corpList.do?method=loadInitPage))

file_path = "F:\\study\\coding\\python\\crawling\\fnguide_financial_statement"   # 종목코드 엑셀파일 경로
file_name = "상장법인목록.xlsx"

company_codes = pd.read_excel(join(file_path, file_name), dtype=str)  # 코스피 종목코드 불러오기

company_data = company_codes[['회사명', '종목코드']]  # 분류, 회사명, 종목코드만 가져오기

# 회사이름 입력하면 그 이름이 있는지와 있다면 종목코드 가져오기
company_name = input('회사명을 입력해주세요 : ')

while len(company_data[company_data['회사명'] == company_name]) == 0:  # 회사이름이 완전히 같지않으면 즉, 없거나 일부포함하면 0을 리턴
    print('해당 이름의 회사가 존재하지 않습니다. 다시 입력해주세요')
    print('아래에 회사목록이 있다면 아래 회사중 하나를 찾으시나요? 다시 입력해주세요.')
    for row in company_data['회사명']:
        if row.find(company_name) != -1:  # 일부포함하는 회사명이 있는지 확인
            print(row)
    company_name = input('회사명을 입력해주세요 : ')

code = company_data[company_data.회사명 == company_name].종목코드.iloc[0]   # 입력한 회사명과 일치하는 종목코드 리턴
print("회사명: "+company_name+"\n종목코드: "+code)


# def fn_crawling(start_point=0, num=0):
num = 0

sheet_row = []
sheet_columns = ['회사명', '시가총액(억)', '배당수익률', 'BPS', '평균ROE', '현재가', '5년후 미래가치', '기대수익률', '투자가능가격', '주가점수', '수익률점수', '점수(20점만점)']
df_snow = pd.DataFrame(columns=sheet_columns)


try:
    # fnguide 로 크롤링 변경예정..
    # snapshot url
    # http://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701
    # 재무제표 url
    # http://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&gicode=A005930&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701
    # 재무비율 url
    # http://comp.fnguide.com/SVO2/ASP/SVD_FinanceRatio.asp?pGB=1&gicode=A005380&cID=&MenuYn=Y&ReportGB=&NewMenuID=104&stkGb=701
    # 투자지표 url
    # http://comp.fnguide.com/SVO2/ASP/SVD_Invest.asp?pGB=1&gicode=A005380&cID=&MenuYn=Y&ReportGB=&NewMenuID=105&stkGb=701
    url = []
    page_list = []
    page_source = []
    tree_list = []
    company_index = {}

    ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36"
    headers = {'User-Agent': ua}

    url.append("http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A"+code+"&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701")
    url.append("http://comp.fnguide.com/SVO2/ASP/SVD_Finance.asp?pGB=1&gicode=A"+code+"&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701")
    url.append("http://comp.fnguide.com/SVO2/ASP/SVD_FinanceRatio.asp?pGB=1&gicode=A"+code+"&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701")
    url.append("http://comp.fnguide.com/SVO2/ASP/SVD_Invest.asp?pGB=1&gicode=A"+code+"&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701")

    url_naver = "https://finance.naver.com/item/sise.nhn?code="+code
    page_naver = requests.get(url_naver, headers=headers)
    tree_naver = html.fromstring(page_naver.text)

    # beautifulSoup 으로 페이지 소스 가져오기
    for i in range(len(url)):
        page_list.append(requests.get(url[i], headers=headers).text)  # Snapshot 재무제표 재무비율 투자지표 페이지 소스 담기 text 또는 content
        page_source.append(BeautifulSoup(page_list[i], "html.parser"))  # 페이지소스 html로 볼수 있게, beautifulsoup는 x path를 지원하지 않는다고함, lmxl로 x paht 가능
        tree_list.append(html.fromstring(page_list[i]))

    # # x path를 이용하기위해 selenium 으로 설정하기
    # path = 'C:\\Users\\zow1\\Downloads\\chromedriver_win32'
    # driver = webdriver.Chrome(executable_path=r'F:\study\coding\python\chromedriver_win32\chromedriver.exe')
    # driver.get(url[0])
    # aaa = driver.find_element_by_xpath("//*[@id='svdMainGrid1']/table/tbody/tr[1]/td[1]/text()")
    # print(aaa.text)

    # x path를 이용해 원하는 값 가져오기, 뒤에는 테이블전체로 부터 가져와야할듯
    # 우선 눈덩이주식투자법만 계산.
    # company_index['종가'] = tree_list[0].xpath("//*[@id='svdMainChartTxt11']")[0].text
    # company_index['52주 최고가'] = tree_list[0].xpath("//*[@id='svdMainGrid1']/table/tbody/tr[2]/td[1]")[0].text.split('/')[0]
    # company_index['52주 최저가'] = tree_list[0].xpath("//*[@id='svdMainGrid1']/table/tbody/tr[2]/td[1]")[0].text.split('/')[1]
    company_index['시가총액'] = tree_list[0].xpath("//*[@id='svdMainGrid1']/table/tbody/tr[5]/td[1]")[0].text
    # company_index['발행주식수'] = tree_list[0].xpath("//*[@id='svdMainGrid1']/table/tbody/tr[7]/td[1]")[0].text.split('/')[0]

    company_index['배당수익률'] = tree_list[0].xpath("//*[@id='corp_group2']/dl[5]/dd")[0].text.split('%')[0]
    if company_index['배당수익률'] == '-':
        company_index['배당수익률'] = company_index['배당수익률'].replace('-', '')

    try:
        company_index['BPS'] = tree_list[3].xpath("//*[@id='p_grid1_5']/td[5]")[0].text
        company_index['ROE-3'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[1]")[0].text
        company_index['ROE-2'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[2]")[0].text  # None 값이면 /span 추가
        company_index['ROE-1'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[3]")[0].text
        company_index['ROE-0'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[4]")[0].text
        if company_index['BPS'] is None:
            company_index['BPS'] = tree_list[0].xpath("//*[@id='p_grid1_5']/td[5]/span")[0].text
        if company_index['ROE-3'] is None:
            company_index['ROE-3'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[1]/span")[0].text
        if company_index['ROE-2'] is None:
            company_index['ROE-2'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[2]/span")[0].text
        if company_index['ROE-1'] is None:
            company_index['ROE-1'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[3]/span")[0].text
        if company_index['ROE-0'] is None:
            company_index['ROE-0'] = tree_list[0].xpath("//*[@id='highlight_D_A']/table/tbody/tr[17]/td[4]/span")[0].text
    except IndexError:
        pass
    company_index['현재가'] = tree_naver.xpath("//*[@id='_nowVal']")[0].text
    # current_price = float(company_index['현재가'].replace(',', ''))
    # total_price = float(company_index['시가총액'].replace(',', ''))

    for key in company_index:
        # company_index[key] = company_index[key].replace(',', '')
        try:
            company_index[key] = float(company_index[key].replace(',', ''))
        except ValueError:
            company_index[key] = None
            continue
    if company_index['ROE-3'] and company_index['ROE-2'] and company_index['ROE-1'] and company_index['ROE-0'] and company_index['BPS'] is not None:

        my_roe = 1.15
        avg_roe = (company_index['ROE-3']+company_index['ROE-2']+company_index['ROE-1']+company_index['ROE-0'])/400
        expected_price = company_index['BPS']*pow(1+avg_roe, 5)
        multiplier = expected_price/company_index['현재가']
        expected_ratio = pow(multiplier, 1/5)
        investible_price = expected_price/(pow(my_roe, 5)) - 1

        recommended_items = []
        if expected_price > 0 and type(expected_ratio) is not complex:

            if expected_ratio > my_roe and investible_price > company_index['현재가']:
                recommended_items.append(company_name)
                recommended_items.append(company_index['시가총액'])
                recommended_items.append(company_index['배당수익률'])
                recommended_items.append(company_index['BPS'])
                recommended_items.append(avg_roe)
                recommended_items.append(company_index['현재가'])
                recommended_items.append(expected_price)
                recommended_items.append(expected_ratio)
                recommended_items.append(investible_price)
                std1 = (company_index['현재가']/investible_price)
                # 현재가가 투자가능가격보다 50%이상 낮으면 10점, 60%면 9점 70%면 8점 80%면 7점 90%면 6점
                # 기대수익률이 25%이상이면 10점 20%이상이면 9점 17,5%이상이면 8점 15%이상이면 7점
                if std1 < 0.5:
                    score = 10
                elif 0.5 < std1 <= 0.6:
                    score = 9
                elif 0.6 < std1 <= 0.7:
                    score = 8
                elif 0.7 < std1 <= 0.8:
                    score = 7
                else:
                    score = 6
                if expected_ratio > 0.25:
                    score2 = 10
                elif expected_ratio > 0.20:
                    score2 = 9
                elif expected_ratio > 0.175:
                    score2 = 8
                else:
                    score2 = 7
                total_score = score + score2
                recommended_items.append(score)
                recommended_items.append(score2)
                recommended_items.append(total_score)
                df_snow.loc[num] = recommended_items
                num += 1
    # start_point += 1
    sleep(1)
except (TimeoutError, requests.exceptions.ConnectionError) as e:
        print(e)
        # writer = pd.ExcelWriter("F:\study\주식\\상장법인_눈덩이주식.xlsx", engine='xlsxwriter')
        # df_snow.to_excel(writer)
    # return start_point, df_snow

writer = pd.ExcelWriter("F:\study\주식\\상장법인_눈덩이주식_개별.xlsx", engine='xlsxwriter')
df_snow.to_excel(writer)

# try:
#     fn_crawling()
#
# except (TimeoutError, requests.exceptions.ConnectionError) as e:
#     print(e)
#     re_start, df = fn_crawling()
#     writer = pd.ExcelWriter("F:\study\주식\\상장법인_눈덩이주식.xlsx", engine='xlsxwriter')
#     df.to_excel(writer)
#     sleep(1)
#     fn_crawling(re_start)



