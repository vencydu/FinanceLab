"""
apply the core-system computation method to calculate the interest
"""

import math
import pandas as pd
import numpy_financial as npf
import datetime
import calendar
import scipy.optimize
from tqdm import tqdm
import numpy
from dateutil.relativedelta import relativedelta
import PySimpleGUI as sg


sg.theme('BrownBlue')
layout= [
    [
        sg.FileBrowse(button_text='请选择输入参数表',
        target = 'INPUT_PATH'), sg.InputText(key= 'INPUT_PATH')
    ],
    [sg.FolderBrowse(button_text='请选择输出文件所在文件夹', target='OUTPUT_FOLDER_PATH'), sg.InputText(key='OUTPUT_FOLDER_PATH')],
    [sg.Text('请选择模型计算内容'), sg.Combo(('Income Only', 'Cost Only', 'BOTH'), default_value='Income Only', key = 'combochoice')],
    [sg.Text('本模型由DKJ开发')],
    [sg.Button('OK'), sg.Button('Cancel')]
]
window = sg.Window('测算表模型', layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == None or event == 'Cancel':
        window.close()
        break
    if event == 'OK':
        sg.popup('已提交将运行') #non_blocking=True)
        window.close()
        break

INPUT_PATH = values['INPUT_PATH']
OUTPUT_PATH = values['OUTPUT_FOLDER_PATH']

FLAG_INCOME = False
FLAG_COST = False
if values['combochoice'] == 'Income Only':
    FLAG_INCOME = True
elif values['combochoice'] == 'Cost Only':
    FLAG_COST = True
elif values['combochoice'] == 'BOTH':
    FLAG_COST = True
    FLAG_INCOME = True

if FLAG_INCOME == True:
    yearly_pattern = pd.read_excel(r'D:\2APPLICATIONS\参数表完整v1.xlsx', sheet_name='月间节奏投放端', index_col=0).reset_index()
    progress_arrangements = pd.read_excel(r'D:\2APPLICATIONS\参数表完整v1.xlsx', sheet_name='月内节奏投放端')
    details = pd.read_excel(r'D:\2APPLICATIONS\参数表完整v1.xlsx', sheet_name='合同参数投放端')
    tax_rate = pd.read_excel(r'D:\2APPLICATIONS\参数表完整v1.xlsx', sheet_name='收入税率投放端', index_col=0).reset_index()
    fee = pd.read_excel(r'D:\2APPLICATIONS\参数表完整v1.xlsx', sheet_name='服务费投放端', index_col=0).reset_index()


    def xnpv(rate, values, dates):
        if rate <= -1.0:
            return float('inf')
        d0 = dates[0]    # or min(dates)
        return sum([ vi / (1.0 + rate)**((di - d0).days / 365.0) for vi, di in zip(values, dates)])

    def xirr(values, dates):
        try:
            return scipy.optimize.newton(lambda r: xnpv(r, values, dates), 0.0)
        except RuntimeError:
            return scipy.optimize.brentq(lambda r: xnpv(r, values, dates), -1.0, 1e10)

    yearly_pattern_transport = yearly_pattern.melt(id_vars='月份', var_name='事业部', value_name='月间起租比例')
    yearly_pattern_transport.columns = ['事业部','月份','月间起租比例']
    yearly_pattern_transport = yearly_pattern_transport[yearly_pattern_transport['月间起租比例']>0] #只选取有起租比例的月份

    progress_arrangements.columns = ['事业部','day','ratio']
    progress = yearly_pattern_transport.merge(progress_arrangements, left_on='事业部', right_on='事业部', how='inner')

    details = details[details['起租金额']>0]
    data = details.merge(progress, right_on='事业部', left_on='事业部', how='left').sort_values(['事业部','产品名称','年份','月份','day'])
    data = data.reset_index().rename(columns={'index':'项目编号'}).sort_values(['事业部','项目编号','产品名称','年份','月份','day'])

    tax_rate_transport = tax_rate.melt(id_vars='事业部',var_name='事业部a',value_name='产品税率')
    tax_rate_transport.columns=['事业部','产品','产品税率']
    data = data.merge(tax_rate_transport, left_on=['事业部','产品名称'], right_on=['事业部','产品'], how='inner')

    data['project_amount'] = data.apply(lambda item: item['起租金额']*item['ratio']*item['月间起租比例'], axis=1)
    data['date'] = data.apply(lambda item: str(item['年份'])+'-'+str(int(item['月份']))+'-'+str(int(item['day'])), axis=1)
    data['fee'] = data.apply(lambda item: item['project_amount']*item['服务费率'], axis=1)
    data['margin'] = data.apply(lambda item: item['project_amount']*item['保证金率'], axis=1)
    data['bank_note'] = data.apply(lambda item: item['project_amount']*item['银票比例'], axis=1)
    data['actual_cash_flow'] = data.apply(lambda item: -item['project_amount']+item['fee']+item['margin']+item['bank_note'], axis=1)
    data['key'] = data.apply(lambda item: item['事业部']+'-'+item['产品名称']+'-'+str(item['date']), axis=1)
    data = data.drop(['产品'], axis=1)

    """
    开始计算FCF
    """

    zero_rate_list = pd.DataFrame()
    final_result = pd.DataFrame()
    for ii in tqdm(range(data.shape[0])):
    #for ii in range (1):
        count = range(data.shape[0])
        sg.one_line_progress_meter('XIRR process', ii+1, len(count), 'XIRR Projection in Progress')
        case_result = []
        invalid_contract_list = []
        current_case = data.iloc[ii, :]
        case_date = datetime.datetime.strptime(current_case['date'], '%Y-%m-%d').date()
        case_total_period = current_case['久期年']*12
        case_total_period_floor = math.floor(case_total_period)
        case_pmt = -npf.pmt(current_case['利率']/int(current_case['年还租次数']),
                           round(case_total_period/12*current_case['年还租次数']),current_case['project_amount'])
        case_bank_note = current_case['bank_note']
        case_margin = current_case['margin']
        case_dpt = current_case['事业部']
        case_date_banknote = case_date+datetime.timedelta(int(current_case['银票期限']))
        case_interval=12.0/current_case['年还租次数']
        case_daily_interest_intoaccount= current_case['project_amount']*current_case['利率']/365
        case_result.append([current_case['公司名称'], current_case['事业部'], current_case['产品名称'],current_case['date'],
                            current_case['project_amount'], current_case['久期年'], current_case['利率'],current_case['项目编号'],
                            case_date, 0 , -current_case['project_amount'], -current_case['project_amount'],
                            current_case['project_amount'],0,
                            current_case['fee'], current_case['margin'],current_case['bank_note'],
                            current_case['actual_cash_flow'], current_case['产品税率'],
                            current_case['key']
                            ])
        case_final_date = case_date + relativedelta(months=case_total_period_floor)
        if case_daily_interest_intoaccount ==0:
            print('contract has 0 interest rate')
            invalid_contract_list.append([
                current_case['公司名称'], current_case['事业部'], current_case['产品名称'], current_case['date'],
                current_case['project_amount'], current_case['久期年'], current_case['利率'], current_case['项目编号']
            ])
            zero_rate_contract = pd.DataFrame(invalid_contract_list)
            zero_rate_list = pd.concat([zero_rate_list, zero_rate_contract])

            continue

        for _ in range(round(case_total_period/12*current_case['年还租次数'])):
            former_date = case_result[-1][8]
            new_date = min(former_date + relativedelta(months=case_interval), case_final_date)
            if case_date_banknote >= former_date and case_date_banknote < new_date and case_bank_note>0:
                new_interest_into_account_banknote = case_result[-1][12]/365*current_case['利率']
                case_result.append(
                    [current_case['公司名称'], current_case['事业部'], current_case['产品名称'], current_case['date'],
                     current_case['project_amount'], current_case['久期年'], current_case['利率'], current_case['项目编号'],
                     case_date_banknote, 0,0,0, case_result[-1][12], new_interest_into_account_banknote, 0,0,
                     -case_bank_note, -case_bank_note, current_case['产品税率'], current_case['key']
                    ]
                )

            if current_case['计息方式']== 'daily':
                daily_date_diff = new_date - former_date
                interest_rate = current_case['利率']*daily_date_diff.days/360.0
                new_interest = case_result[-1][12]*interest_rate
            elif current_case['计息方式']== 'monthly':
                monthly_date_diff = (new_date.year - former_date.year)*12+(new_date.month - former_date.month)
                interest_rate = current_case['利率']/12*monthly_date_diff
                new_interest = case_result[-1][12]*interest_rate

            if new_date== case_final_date:
                new_interest = case_result[-1][12]*current_case['利率']*(new_date-former_date).days/360.0
            if new_date == case_final_date:
                new_interest_into_account_daily= 0

            new_principle = case_pmt - new_interest #每期的本金=每期pmt租金-每期利息

            new_uncollected_principal = case_result[-1][12] - new_principle
            new_interest_into_account_daily = case_result[-1][12]/365*current_case['利率']

            case_result.append([current_case['公司名称'], current_case['事业部'], current_case['产品名称'], current_case['date'],
                     current_case['project_amount'], current_case['久期年'], current_case['利率'], current_case['项目编号'],
                     new_date, new_interest, new_principle, case_pmt, new_uncollected_principal,
                     new_interest_into_account_daily, 0,0, 0, case_pmt, current_case['产品税率'], current_case['key']

            ])

        case_result[-1][10] += case_result[-1][12]
        case_result[-1][11] += case_result[-1][12]
        case_result[-1][17] += case_result[-1][12]
        case_result[-1][12] = 0

        case_result[-1][17] -= case_margin
        case_result[-1][15] = -case_margin
        FCF_result = pd.DataFrame(case_result)
        FCF_result = FCF_result.sort_values(8)
        FCF_result.columns = ['公司名称','事业部','产品名称','起租日期','项目金额','项目年限','利率','项目编号','date','本期利息','本期还款本金',
                              'pmt本期租金','总未回收本金','上期当日利息','服务费','保证金','银票','自由现金流','产品税率','key'
                            ]
        FCF_result['XIRR'] = xirr(FCF_result['自由现金流'], FCF_result['date'])
        final_result =pd.concat([final_result, FCF_result])

    if zero_rate_list.empty == False:
        zero_rate_list.columns = ['公司名称','事业部','产品名称','date','project_amount', '久期年','利率','项目编号']

    final_result['date'] = final_result['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d'))
    final_result_data = final_result.drop_duplicates(subset=['项目编号']).sort_values(['项目编号'])
    final_result_exle = final_result.iloc[:]
    final_result_exle['date'] = final_result_exle['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())


    """
    计算XIRR
    """
    income_XIRR_data = final_result_data.iloc[:]
    income_XIRR_data['权数'] = income_XIRR_data['项目金额'] * income_XIRR_data['项目年限']
    income_XIRR_data['权数*XIRR'] = income_XIRR_data['权数'] * income_XIRR_data['XIRR']
    income_XIRR_data['权数*利率'] = income_XIRR_data['权数'] * income_XIRR_data['利率']

    income_XIRR_data_exle = income_XIRR_data.iloc[:]
    income_XIRR_data_exle['date'] = income_XIRR_data_exle['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())

    """计算XIRR公司年度数据"""
    income_XIRR_year_1 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称', '事业部', '产品名称'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_year_1['XIRR'] = income_XIRR_year_1['权数*XIRR'] / income_XIRR_year_1['权数']
    income_XIRR_year_1['利率'] = income_XIRR_year_1['权数*利率'] / income_XIRR_year_1['权数']
    income_XIRR_year_1['date'] = income_XIRR_year_1['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_year_1 = income_XIRR_year_1.rename(columns = {'XIRR': '累计XIRR', '利率': '累计利率'})
    income_XIRR_pivot_yr1 = income_XIRR_year_1.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['累计XIRR', '累计利率']).reset_index()

    income_XIRR_year_2 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称', '事业部'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_year_2['XIRR'] = income_XIRR_year_2['权数*XIRR'] / income_XIRR_year_2['权数']
    income_XIRR_year_2['利率'] = income_XIRR_year_2['权数*利率'] / income_XIRR_year_2['权数']
    income_XIRR_year_2['date'] = income_XIRR_year_2['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_year_2 = income_XIRR_year_2.rename(columns = {'XIRR': '累计XIRR', '利率': '累计利率'})
    income_XIRR_year_2['产品名称'] = '小计'
    income_XIRR_pivot_yr2 = income_XIRR_year_2.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['累计XIRR', '累计利率']).reset_index()

    income_XIRR_year_3 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_year_3['XIRR'] = income_XIRR_year_3['权数*XIRR'] / income_XIRR_year_3['权数']
    income_XIRR_year_3['利率'] = income_XIRR_year_3['权数*利率'] / income_XIRR_year_3['权数']
    income_XIRR_year_3['date'] = income_XIRR_year_3['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_year_3 = income_XIRR_year_3.rename(columns = {'XIRR': '累计XIRR', '利率': '累计利率'})
    income_XIRR_year_3['产品名称'] = '总计'
    income_XIRR_year_3['事业部'] = '总计'
    income_XIRR_pivot_yr3 = income_XIRR_year_3.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['累计XIRR', '累计利率']).reset_index()

    income_XIRR_pivot_yr = pd.concat([income_XIRR_pivot_yr1, income_XIRR_pivot_yr2, income_XIRR_pivot_yr3])


    """计算XIRR公司月度数据"""
    income_XIRR_month_1 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称', '事业部', '产品名称'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_month_1['XIRR'] = income_XIRR_month_1['权数*XIRR'] / income_XIRR_month_1['权数']
    income_XIRR_month_1['利率'] = income_XIRR_month_1['权数*利率'] / income_XIRR_month_1['权数']
    income_XIRR_month_1['date'] = income_XIRR_month_1['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_pivot_m1 = income_XIRR_month_1.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['XIRR', '利率']).reset_index()

    income_XIRR_month_2 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称', '事业部'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_month_2['XIRR'] = income_XIRR_month_2['权数*XIRR'] / income_XIRR_month_2['权数']
    income_XIRR_month_2['利率'] = income_XIRR_month_2['权数*利率'] / income_XIRR_month_2['权数']
    income_XIRR_month_2['产品名称'] = '小计'
    income_XIRR_month_2['date'] = income_XIRR_month_2['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_pivot_m2 = income_XIRR_month_2.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['XIRR', '利率']).reset_index()

    income_XIRR_month_3 = income_XIRR_data.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称'])[['权数', '权数*XIRR', '权数*利率']].sum().reset_index()
    income_XIRR_month_3['XIRR'] = income_XIRR_month_3['权数*XIRR'] / income_XIRR_month_3['权数']
    income_XIRR_month_3['利率'] = income_XIRR_month_3['权数*利率'] / income_XIRR_month_3['权数']
    income_XIRR_month_3['产品名称'] = '总计'
    income_XIRR_month_3['事业部'] = '总计'
    income_XIRR_month_3['date'] = income_XIRR_month_3['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_XIRR_pivot_m3 = income_XIRR_month_3.pivot_table(index = ['公司名称','事业部','产品名称'], columns=['date'], values= ['XIRR', '利率']).reset_index()


    income_XIRR_pivot_m = pd.concat([income_XIRR_pivot_m1, income_XIRR_pivot_m2, income_XIRR_pivot_m3])

    """链接结果"""
    income_XIRR_pivot = income_XIRR_pivot_yr.merge(income_XIRR_pivot_m, on = ['公司名称','事业部','产品名称'])


    """
    计算资产余额
    """
    """年底资产余额年的维度"""
    income_asset_yr1 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称', '事业部', '产品名称'])['本期还款本金'].sum().reset_index()
    income_asset_yr_cum1 = income_asset_yr1.groupby(['公司名称', '事业部', '产品名称', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_yr_cum1['date'] = income_asset_yr_cum1['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_yr_cum1 = income_asset_yr_cum1.rename(columns= {'本期还款本金':'年末资产余额'})
    income_asset_yr_cum1['年末资产余额'] = income_asset_yr_cum1['年末资产余额'].apply(lambda item: 0 - item)
    income_asset_pivot_yr1 = income_asset_yr_cum1.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['年末资产余额']).reset_index()

    income_asset_yr2 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称', '事业部'])['本期还款本金'].sum().reset_index()
    income_asset_yr_cum2 = income_asset_yr2.groupby(['公司名称', '事业部', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_yr_cum2['date'] = income_asset_yr_cum2['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_yr_cum2 = income_asset_yr_cum2.rename(columns= {'本期还款本金':'年末资产余额'})
    income_asset_yr_cum2['产品名称'] = '小计'
    income_asset_yr_cum2['年末资产余额'] = income_asset_yr_cum2['年末资产余额'].apply(lambda item: 0 - item)
    income_asset_pivot_yr2 = income_asset_yr_cum2.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['年末资产余额']).reset_index()

    income_asset_yr3 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='Y'), '公司名称'])['本期还款本金'].sum().reset_index()
    income_asset_yr_cum3 = income_asset_yr3.groupby(['公司名称', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_yr_cum3['date'] = income_asset_yr_cum3['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_yr_cum3 = income_asset_yr_cum3.rename(columns= {'本期还款本金':'年末资产余额'})
    income_asset_yr_cum3['产品名称'] = '总计'
    income_asset_yr_cum3['事业部'] = '总计'
    income_asset_yr_cum3['年末资产余额'] = income_asset_yr_cum3['年末资产余额'].apply(lambda item: 0 - item)
    income_asset_pivot_yr3 = income_asset_yr_cum3.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['年末资产余额']).reset_index()

    income_asset_pivot_yr = pd.concat([income_asset_pivot_yr1, income_asset_pivot_yr2, income_asset_pivot_yr3])

    """年度资产月的维度"""
    income_asset_m1 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称', '事业部', '产品名称'])['本期还款本金'].sum().reset_index()
    income_asset_m_cum1 = income_asset_m1.groupby(['公司名称', '事业部', '产品名称', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_m_cum1['date'] = income_asset_m_cum1['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_m_cum1 = income_asset_m_cum1.rename(columns= {'本期还款本金':'资产余额'})
    income_asset_m_cum1['资产余额'] = income_asset_m_cum1['资产余额'].apply(lambda item: 0 - item)
    income_asset_pivot_m1 = income_asset_m_cum1.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['资产余额']).reset_index()

    income_asset_m2 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称', '事业部'])['本期还款本金'].sum().reset_index()
    income_asset_m_cum2 = income_asset_m2.groupby(['公司名称', '事业部', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_m_cum2['date'] = income_asset_m_cum2['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_m_cum2 = income_asset_m_cum2.rename(columns= {'本期还款本金':'资产余额'})
    income_asset_m_cum2['资产余额'] = income_asset_m_cum2['资产余额'].apply(lambda item: 0 - item)
    income_asset_m_cum2['产品名称'] = '小计'
    income_asset_pivot_m2 = income_asset_m_cum2.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['资产余额']).reset_index()

    income_asset_m3 = final_result.groupby([pd.Grouper(key='date', axis=0, freq='M'), '公司名称'])['本期还款本金'].sum().reset_index()
    income_asset_m_cum3 = income_asset_m3.groupby(['公司名称', 'date']).sum().groupby(level=0).cumsum().reset_index()
    income_asset_m_cum3['date'] = income_asset_m_cum3['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_asset_m_cum3 = income_asset_m_cum3.rename(columns= {'本期还款本金':'资产余额'})
    income_asset_m_cum3['资产余额'] = income_asset_m_cum3['资产余额'].apply(lambda item: 0 - item)
    income_asset_m_cum3['产品名称'] = '总计'
    income_asset_m_cum3['事业部'] = '总计'
    income_asset_pivot_m3 = income_asset_m_cum3.pivot_table(index=['公司名称', '事业部', '产品名称'], columns=['date'], values=['资产余额']).reset_index()

    income_asset_pivot_m = pd.concat([income_asset_pivot_m1, income_asset_pivot_m2, income_asset_pivot_m3])
    """合并结果"""
    income_asset_pivot = income_asset_pivot_yr.merge(income_asset_pivot_m, on = ['公司名称', '事业部', '产品名称'])

    """
    计算利息收入分摊（财务新核心口径）
    """

    cf = final_result
    cf = cf.iloc[:]
    cf['date'] = cf['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    cf['起租日期'] = cf['起租日期'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d').date())
    cf = cf[~((cf['date'] != cf['起租日期']) & (cf['本期利息'] == 0))]
    oneline_contract = pd.DataFrame()
    final_result_income = pd.DataFrame()

    count_income = cf['项目编号'].drop_duplicates().shape[0]

    for jj, index in tqdm(enumerate(cf['项目编号'].drop_duplicates())):
        sg.one_line_progress_meter('Income Process', jj + 1, count_income, 'Income Projection in Progress')
        invalid_income_contract_list = []
        current_info = cf[cf['项目编号'] == index]
        if current_info.shape[0] == 1:
            print('contract only has 1 line')
            invalid_income_contract_list.append(index)
            oneline_income = pd.DataFrame(invalid_income_contract_list)
            oneline_contract = pd.concat([oneline_contract, oneline_income])
            continue
        current_info_date = current_info['起租日期'].iloc[0]

        former_interest_date = current_info['date'].iloc[0]
        case_result = []
        current_residual_days = 0
        former_case_interest = 0
        former_diff_in_day = 0
        current_case_interest = 0
        current_case = None
        for ii in range(current_info.shape[0]):
            #对每个利息区间进行计算
            #current_case = current_info.loc[current_info['date'] == interest_date, :]
            current_case = current_info.iloc[ii, :]
            current_case_interest = current_case['本期利息']
            current_case_interest_date = current_case['date']
            current_residual_days = current_case_interest_date.day-1
            diff_in_month = (current_case_interest_date.year - former_interest_date.year) * 12 + (current_case_interest_date.month - former_interest_date.month)
            diff_in_day = (current_case_interest_date - former_interest_date).days
            former_booking_date = former_interest_date
            test_sum = 0
            for month_cnt in range(diff_in_month):
                #生成每个记账周期的记录
                #booking_date = former_interest_date.replace(day = calendar.monthrange(former_interest_date.year, former_interest_date.month)[1]) + relativedelta.relativedelta(months=month_cnt)
                # 对合同的第一期进行调整
                if month_cnt == 0:
                    former_booking_date1 = former_booking_date + datetime.timedelta(days=-32)
                else:
                    former_booking_date1 = former_booking_date
                former_booking_date2 = former_booking_date1.replace(day=28) + datetime.timedelta(days=4)
                former_booking_date3 = former_booking_date2.replace(day=28) + datetime.timedelta(days=4)
                booking_date = former_booking_date3 - datetime.timedelta(days=former_booking_date3.day)
                days = (booking_date - former_booking_date).days

                if month_cnt == 0:
                    days += 1
                booking_interest = days * current_case['本期利息'] / diff_in_day

                if ii != 1 and month_cnt == 0:
                    booking_interest += current_residual_days * former_case_interest / former_diff_in_day
                # 对还利息的第一期进行调整
                case_result.append(list(current_case[['公司名称', '事业部', '产品名称', '产品税率', '起租日期', '项目金额', '项目年限', '利率', '项目编号', '本期利息']]) + [former_interest_date, booking_date, booking_interest])
                test_sum += booking_interest
                #print(ii, test_sum, month_cnt, booking_date, former_booking_date, days, current_residual_days * former_case_interest / former_diff_in_day, booking_interest, current_case['本期利息'], diff_in_day)
                former_booking_date = booking_date
            former_case_interest = current_case_interest
            former_diff_in_day = diff_in_day
            former_interest_date = current_case['date']
        # 最后一期
        booking_interest = current_residual_days * former_case_interest / former_diff_in_day
        booking_date = former_interest_date.replace(day = calendar.monthrange(former_interest_date.year, former_interest_date.month)[1])
        case_result.append(list(current_case[['公司名称', '事业部', '产品名称', '产品税率', '起租日期', '项目金额', '项目年限', '利率', '项目编号', '本期利息']]) + [former_interest_date, booking_date, booking_interest])
       # print(1000, month_cnt, booking_date, former_booking_date, days, booking_interest, current_case['本期利息'], diff_in_day)

        case_result_pd = pd.DataFrame(case_result, columns = ['公司名称', '事业部', '产品名称', '产品税率', '起租日期', '项目金额', '项目年限', '利率', '项目编号', '本期利息', 'date', 'booking_date', 'booking_interest'])
        final_result_income = pd.concat([final_result_income, case_result_pd])

    final_result_income = final_result_income.rename(columns = {'booking_date' : '入账日期', 'booking_interest' : '利息收入'})
    final_result_income['税后利息收入'] = final_result_income.apply(lambda item: item['利息收入'] / (1 + item['产品税率']), axis=1)

    if oneline_contract.empty == False:
        oneline_contract.columns = ['项目编号']


    """服务费分摊"""
    data['date'] = pd.to_datetime(data['date'], format = '%Y-%m-%d')
    data_pivot = data.groupby([pd.Grouper(key = 'date', axis=0, freq='M'), '公司名称', '事业部', '产品名称'])['fee'].sum().reset_index()
    data_pivot = data_pivot.reset_index().rename(columns={'index':'服务费编号'}).sort_values(['公司名称', '事业部', '产品名称'])

    fee_transport = fee.melt(id_vars='事业部' ,var_name='分摊方式', value_name='占比')
    fee_data = data_pivot.merge(fee_transport, left_on = '事业部', right_on = '事业部')
    fee_data['fee_tballocate'] = fee_data.apply(lambda item: item['fee'] * item['占比'], axis=1)
    fee_allo = {'按照一个月分摊':'1', '按照两个月分摊':'2', '按照三个月分摊':'3', '按照四个月分摊':'4' , '按照实际利率法分摊':'30'}
    fee_data['分摊时长'] = fee_data['分摊方式'].apply(lambda  item: fee_allo[item])

    final_result_fee = pd.DataFrame()
    for ii in tqdm(range(fee_data.shape[0])):
        count_fee = range(fee_data.shape[0])
        sg.one_line_progress_meter('process', ii+1, len(count_fee), 'Fee Projection in Progress')
        case_result = []
        current_case = fee_data.iloc[ii, :]
        case_date = current_case['date']
        case_period = current_case['分摊时长']
        case_fee_term = current_case['fee_tballocate'] / int(case_period)
        case_final_date = case_date + relativedelta(months=int(case_period) - 1)
        case_result.append([current_case['服务费编号'], current_case['公司名称'], current_case['事业部'],
                            current_case['产品名称'], current_case['date'], current_case['fee'],
                            current_case['分摊方式'], current_case['fee_tballocate'], case_date, case_fee_term])

        for _ in range(int(case_period) - 1):
            former_date = case_result[-1][8]
            new_date = former_date + relativedelta(months=1)

            case_result.append([current_case['服务费编号'], current_case['公司名称'], current_case['事业部'],
                            current_case['产品名称'], current_case['date'], current_case['fee'],
                            current_case['分摊方式'], current_case['fee_tballocate'], new_date, case_fee_term])

        FEE_result = pd.DataFrame(case_result)
        FEE_result = FEE_result.sort_values(8)
        FEE_result.columns = ['服务费编号', '公司名称', '事业部', '产品名称', 'date', 'fee', '分摊方式', '待分摊金额', '入账日期', '分摊服务费税前']
        final_result_fee = pd.concat([final_result_fee, FEE_result])

    """处理服务费，公司-事业部-产品-日期的金额具有唯一性"""
    fee_tax_rate_infor = details[['公司名称', '事业部', '产品名称', '服务费税率']]
    final_result_fee = final_result_fee.merge(fee_tax_rate_infor, on = ['公司名称', '事业部', '产品名称'], how = 'left')
    final_result_fee['税后服务费'] = final_result_fee.apply(lambda item: item['分摊服务费税前'] / (1 + item['服务费税率']), axis=1)

    final_result_fee_exle = final_result_fee.iloc[:]
    final_result_fee_exle['date'] = final_result_fee_exle['date'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    final_result_fee_exle['入账日期'] = final_result_fee_exle['入账日期'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())

    fee_pivot = final_result_fee.groupby([pd.Grouper(key='入账日期', axis=0, freq='M'), '公司名称', '事业部', '产品名称'])[['税后服务费']].sum().reset_index()
    """处理利息收入，公司-事业部-产品-日期的金额具有唯一性"""
    final_result_income['入账日期'] = pd.to_datetime(final_result_income['入账日期'], format='%Y-%m-%d')

    final_result_income_exle = final_result_income.iloc[:]
    final_result_income_exle['入账日期'] = final_result_income_exle['入账日期'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())

    interest_pivot = final_result_income.groupby([pd.Grouper(key='入账日期', axis=0, freq='M'),  '公司名称', '事业部', '产品名称'])[['税后利息收入']].sum().reset_index()

    if max(fee_pivot['入账日期']) > max(interest_pivot['入账日期']):
        income_trans = fee_pivot.merge(interest_pivot, on = ['入账日期','公司名称', '事业部', '产品名称'], how='left')
    else:
        income_trans = fee_pivot.merge(interest_pivot, on = ['入账日期','公司名称', '事业部', '产品名称'], how='right')

    income_trans = income_trans.rename(columns={'入账日期':'日期', 'level_4':'收入类型', 0:'金额'})

    income_trans_exle = income_trans.iloc[:]
    income_trans_exle['日期'] = income_trans_exle['日期'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())

    income_PL_data = income_trans.set_index(['日期','公司名称', '事业部', '产品名称']).stack().reset_index()
    income_PL_data = income_PL_data.rename(columns={'level_4':'收入类型', 0:'金额'})


    """收入端结果-年"""
    income_PL_yr1 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='Y'),'公司名称', '收入类型', '事业部', '产品名称'])['金额'].sum().reset_index()
    income_PL_yr1['日期'] = income_PL_yr1['日期'].apply(lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_yr1 = income_PL_yr1.rename(columns={'金额': '年度累计收入'})
    income_PL_yr1_pivot = income_PL_yr1.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'], columns=['日期'], values=['年度累计收入']).reset_index()

    income_PL_yr2 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='Y'), '公司名称', '收入类型', '事业部'])['金额'].sum().reset_index()
    income_PL_yr2['日期'] = income_PL_yr2['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_yr2 = income_PL_yr2.rename(columns={'金额': '年度累计收入'})
    income_PL_yr2['产品名称'] = '小计'
    income_PL_yr2_pivot = income_PL_yr2.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                    columns=['日期'], values=['年度累计收入']).reset_index()

    income_PL_yr3 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='Y'), '公司名称', '收入类型'])['金额'].sum().reset_index()
    income_PL_yr3['日期'] = income_PL_yr3['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_yr3 = income_PL_yr3.rename(columns={'金额': '年度累计收入'})
    income_PL_yr3['产品名称'] = '合计'
    income_PL_yr3['事业部'] = '合计'
    income_PL_yr3_pivot = income_PL_yr3.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                    columns=['日期'], values=['年度累计收入']).reset_index()

    income_PL_yr4 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='Y'), '公司名称'])[
        '金额'].sum().reset_index()
    income_PL_yr4['日期'] = income_PL_yr4['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_yr4 = income_PL_yr4.rename(columns={'金额': '年度累计收入'})
    income_PL_yr4['产品名称'] = '公司总计'
    income_PL_yr4['事业部'] = '公司总计'
    income_PL_yr4['收入类型'] = '公司总计'
    income_PL_yr4_pivot = income_PL_yr4.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                    columns=['日期'], values=['年度累计收入']).reset_index()

    income_PL_pivot_yr = pd.concat([income_PL_yr1_pivot, income_PL_yr2_pivot, income_PL_yr3_pivot, income_PL_yr4_pivot])

    """收入端结果-月"""
    income_PL_m1 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='M'), '公司名称', '收入类型', '事业部', '产品名称'])[
        '金额'].sum().reset_index()
    income_PL_m1['日期'] = income_PL_m1['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_m1 = income_PL_m1.rename(columns={'金额': '月度收入'})
    income_PL_m1_pivot = income_PL_m1.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                    columns=['日期'], values=['月度收入']).reset_index()

    income_PL_m2 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='M'), '公司名称', '收入类型', '事业部'])[
        '金额'].sum().reset_index()
    income_PL_m2['日期'] = income_PL_m2['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_m2 = income_PL_m2.rename(columns={'金额': '月度收入'})
    income_PL_m2['产品名称'] = '小计'
    income_PL_m2_pivot = income_PL_m2.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                  columns=['日期'], values=['月度收入']).reset_index()

    income_PL_m3 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='M'), '公司名称', '收入类型'])[
        '金额'].sum().reset_index()
    income_PL_m3['日期'] = income_PL_m3['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_m3 = income_PL_m3.rename(columns={'金额': '月度收入'})
    income_PL_m3['产品名称'] = '合计'
    income_PL_m3['事业部'] = '合计'
    income_PL_m3_pivot = income_PL_m3.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                  columns=['日期'], values=['月度收入']).reset_index()

    income_PL_m4 = income_PL_data.groupby([pd.Grouper(key='日期', axis=0, freq='M'), '公司名称'])[
        '金额'].sum().reset_index()
    income_PL_m4['日期'] = income_PL_m4['日期'].apply(
        lambda item: datetime.datetime.strptime(str(item), '%Y-%m-%d %H:%M:%S').date())
    income_PL_m4 = income_PL_m4.rename(columns={'金额': '月度收入'})
    income_PL_m4['产品名称'] = '公司总计'
    income_PL_m4['事业部'] = '公司总计'
    income_PL_m4['收入类型'] = '公司总计'
    income_PL_m4_pivot = income_PL_m4.pivot_table(index=['公司名称', '收入类型', '事业部', '产品名称'],
                                                  columns=['日期'], values=['月度收入']).reset_index()

    income_PL_pivot_m = pd.concat([income_PL_m1_pivot, income_PL_m2_pivot, income_PL_m3_pivot, income_PL_m4_pivot])
    """合并结果"""
    income_PL_pivot = income_PL_pivot_yr.merge(income_PL_pivot_m, on = ['公司名称', '收入类型', '事业部', '产品名称'])

    if FLAG_COST == False:
        sg.popup('Wait for excel writing (It may take sometime'
                 '\nWhen finish, another popup resent', non_blocking=True)

        with pd.ExcelWriter(OUTPUT_PATH + '/Forecast投资端结果'+ datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')+ '.xlsx', datetime_format='YYYY-MM-DD') as writer:
            zero_rate_list.to_excel(writer, sheet_name='ERROR投资端零利率合同', index=())
            final_result_exle.to_excel(writer, sheet_name='投放端FCF', index=())
            income_XIRR_data_exle.to_excel(writer, sheet_name='投放项目信息', index=())
            income_XIRR_pivot.to_excel(writer, sheet_name='投放XIRR分事业部')
            income_asset_pivot.to_excel(writer, sheet_name='资产透视结果')
            income_trans_exle.to_excel(writer, sheet_name='收入流水明细', index=())
            final_result_fee_exle.to_excel(writer, sheet_name='服务费流水明细', index=())
            final_result_income_exle.to_excel(writer, sheet_name='利息流水明细', index=())
            oneline_contract.to_excel(writer, sheet_name='ERROR收入单行合同', index=())
            income_PL_pivot.to_excel(writer, sheet_name='利息收入透视后明细')

        sg.popup_ok('INCOME Task Done')