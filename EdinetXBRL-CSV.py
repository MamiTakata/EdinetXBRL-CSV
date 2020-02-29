# -*- coding: utf-8 -*-

from arelle import ModelManager
from arelle import Cntlr
import os
import zipfile
import glob
import pandas as pd
import const        #const.py
from collections import OrderedDict

const.FOLDER_CURRENT = os.path.dirname(os.path.abspath(__file__))

const.FOLDER_EDINET = os.path.join(const.FOLDER_CURRENT,'edinet')
const.FILE_EDINETCODE = os.path.join(const.FOLDER_EDINET,'EdinetcodeDlInfo.csv')
const.FOLDER_XBRL = os.path.join(const.FOLDER_EDINET,'XBRL/PublicDoc/')
const.FILE_CSV = os.path.join(const.FOLDER_EDINET,'CompanyInfo.csv')
const.FILE_EDINET_CSV = os.path.join(const.FOLDER_EDINET,'Edinet.csv')

dict_cols = OrderedDict()
dict_cols = {
     '企業名':
        { 'element_id':['FilerNameInJapaneseDEI'] 
         ,'contextRef':'FilingDateInstant'
        }
    ,'EDINETコード':
        { 'element_id':['EDINETCodeDEI'] 
         ,'contextRef':'FilingDateInstant'
        }
    ,'提出者業種':
        { 'element_id':[] 
         ,'contextRef':''
        }
    ,'単独/連結':
        { 'element_id':[] 
         ,'contextRef':''
        }
    ,'証券コード':
        { 'element_id':['SecurityCodeDEI'] 
         ,'contextRef':'FilingDateInstant'
        }
    ,'当事業年度開始日':
        { 'element_id':['CurrentFiscalYearStartDateDEI'] 
         ,'contextRef':'FilingDateInstant'
        }
    ,'当事業年度終了日':
        { 'element_id':['CurrentFiscalYearEndDateDEI'] 
         ,'contextRef':'FilingDateInstant'
        }
    ,'平均年間給与（円）':
        { 'element_id':['AverageAnnualSalaryInformationAboutReportingCompanyInformationAboutEmployees'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'平均勤続年数（年）':
        { 'element_id':['AverageLengthOfServiceYearsInformationAboutReportingCompanyInformationAboutEmployees'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'平均年齢（歳）':
        { 'element_id':['AverageAgeYearsInformationAboutReportingCompanyInformationAboutEmployees'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'従業員数（人）':
        { 'element_id':['NumberOfEmployees'] 
         ,'contextRef':'CurrentYearInstant'
        }
    , '売上高':
        {'element_id':
                    # 売上高
                    ['RevenuesFromExternalCustomers'
                    ,'TransactionsWithOtherSegments'
                    ,'NetSales'
                    ,'OperatingRevenue1'
                    ,'OperatingRevenue2'
                    ,'GrossOperatingRevenue'
                    ,'OrdinaryIncomeBNK'
                    ,'OperatingIncomeINS'
                    # 営業収益
                    ,'OperatingRevenue1SummaryOfBusinessResults'
                    ,'OperatingRevenue2SummaryOfBusinessResults'
                    # 経常収益
                    ,'OrdinaryIncomeSummaryOfBusinessResults'
                    ]
         ,'contextRef':'CurrentYearDuration'
        }
    , '売上原価':
        {'element_id':
                    # 売上原価
                    ['CostOfSales'
                    # 営業費用
                    ,'OperatingExpenses'
                    # 営業原価
                    ,'OperatingCost'
                    ]
         ,'contextRef':'CurrentYearDuration'
        }
    ,'売上総利益':
        { 'element_id':['GrossProfit'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'営業利益':
        { 'element_id':['OperatingIncome'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'経常利益':
        { 'element_id':['OrdinaryIncome'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'当期純利益':
        { 'element_id':['ProfitLoss'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'親会社の所有者に帰属する利益':
        { 'element_id':['ProfitLossAttributableToOwnersOfParent'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'自己資本比率':
        { 'element_id':['EquityToAssetRatioSummaryOfBusinessResults'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'ROE(自己資本利益率)':
        { 'element_id':['RateOfReturnOnEquitySummaryOfBusinessResults'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'EPS(一株当たり利益)':
        { 'element_id':['BasicEarningsLossPerShareSummaryOfBusinessResults'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'BPS(一株当たり純資産)':
        { 'element_id':['NetAssetsPerShareSummaryOfBusinessResults'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'流動資産':
        { 'element_id':['CurrentAssets'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'固定資産':
        { 'element_id':['NoncurrentAssets'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'資産計':
        {'element_id':
                    ['Assets'
                    ,'TotalAssetsSummaryOfBusinessResults'
                    ,'TotalAssetsIFRSSummaryOfBusinessResults'
                    ]
         ,'contextRef':'CurrentYearInstant'
        }
    ,'流動負債':
        { 'element_id':['CurrentLiabilities'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'固定負債':
        { 'element_id':['NoncurrentLiabilities'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'負債計':
        { 'element_id':['Liabilities'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'純資産':
        { 'element_id':['NetAssetsSummaryOfBusinessResults'] 
         ,'contextRef':'CurrentYearInstant'
        }
    ,'営業CF':
        { 'element_id':['NetCashProvidedByUsedInOperatingActivities'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'投資CF':
        { 'element_id':['NetCashProvidedByUsedInInvestmentActivities'] 
         ,'contextRef':'CurrentYearDuration'
        }
    ,'財務CF':
        { 'element_id':['NetCashProvidedByUsedInFinancingActivities'] 
         ,'contextRef':'CurrentYearDuration'
        }
    }


def load_edinet():

    edinet_list = []

    # xbrl 読み込み
    xbrl_files = glob.glob(const.FOLDER_XBRL + '*.xbrl' )

    # EdinetcodeDlInfo.csv 読み込み
    edinetcodedata = pd.read_csv(const.FILE_EDINETCODE, skiprows=1,encoding='cp932')
 
    for index, xbrl_file in enumerate(xbrl_files):
        ctrl = Cntlr.Cntlr()
        model_manager = ModelManager.initialize(ctrl)
        model_xbrl = model_manager.load(xbrl_file)

        print(xbrl_file, ":", index + 1, "/", len(xbrl_files))

        # XBRLをDataFrameにセット
        factData = pd.DataFrame(
                        data=[(
                            fact.concept.qname.localName, 
                            fact.value, 
                            fact.isNumeric, 
                            fact.contextID,
                            fact.decimals
                        ) for fact in model_xbrl.facts ],
                        columns=[
                            'element_id',
                            'value',
                            'isNumeric',
                            'contextID',
                            'decimals'
                        ]) 
        
        # 取得したい要素
        datalist = []
        datalist_ren = []
        flg_ren = False
        for key, value in dict_cols.items():
            contextRef = value['contextRef']
            element_ids = value['element_id']
            if len(element_ids) > 0:
                for element_id in element_ids :
                    # data1 = factData[factData['element_id'].str.contains(element_id)]
                    data1 = []
                    data2 = []
                    data3 = []
                    data1 = factData[factData['element_id'] == element_id]
                    if len(data1) == 1 : break

                    data2 = data1[data1['contextID'] == (contextRef + '_NonConsolidatedMember')]
                    data3 = data1[data1['contextID'] == contextRef ]
                    if len(data2) >= 1 or len(data3) >= 1 : break

                data,data_ren,unitRef = '','',''
                if len(data1) == 1 :
                    data = data1['value'].values[0]
                    data_ren = data
                    print(key + ' : ' + data)
                if len(data2) >= 1 and len(data3) == 0:
                    data = data2['value'].values[0]
                    data_ren = data
                    print(key + ' : ' + data)
                if len(data3) >= 1 :
                    data = ''
                    data_ren = data3['value'].values[0]
                    print(key + '(連結) : ' + data_ren)
                if len(data1) == 0:
                    print(key + ' : ' + 'NoData ')
                    data = ''

                datalist.append(data)
                if flg_ren :
                    datalist_ren.append(data_ren)
                
                if key == 'EDINETコード':
                    company = edinetcodedata[edinetcodedata['ＥＤＩＮＥＴコード'] == data]
                    if len(company) == 1 :
                        print('提出者業種 : ' + company['提出者業種'].values[0])
                        data = company['提出者業種'].values[0]

                        flg_ren = False
                        if company['連結の有無'].values[0] == '有':
                            print('連結の有無 : ' + company['連結の有無'].values[0] )
                            flg_ren = True
                    else:
                        print('提出者業種 : ' + 'NoData ')
                        data = ''
                    datalist.append(data)
                    datalist.append('連結' if flg_ren else '単独')

                    if flg_ren :
                        # 連結用にデータをコピー
                        datalist_ren = datalist[:]

        if flg_ren :
            edinet_list.append(datalist_ren)
        else:
            edinet_list.append(datalist)

    return edinet_list



def write_edinet_csv(edinet_list):
    df = pd.DataFrame(edinet_list,columns=dict_cols.keys())
    df.to_csv(const.FILE_EDINET_CSV, encoding='cp932')
    print('Finish. Write to csv.')


def unzip_file(zip_dir):
    zip_files = glob.glob(os.path.join(zip_dir, '*.zip'))

    number_of_zip_lists = len(zip_files)
    print("number_of_zip_lists：", number_of_zip_lists)

    for index, zip_file in enumerate(zip_files):
        print(zip_file, ":", index + 1, "/", number_of_zip_lists)
        with zipfile.ZipFile(zip_file) as zip_f:
            zip_f.extractall(zip_dir)
            zip_f.close()


def main():

    # unzip_file(const.FOLDER_EDINET)

    # edinet_list = make_edinet_info()
    # write_csv(edinet_list)

    edinet_list = load_edinet()
    write_edinet_csv(edinet_list)
    print("finish")


if __name__ == "__main__":
    main()
