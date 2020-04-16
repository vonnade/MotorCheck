from xlutils.copy import copy
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext as ScrolledText
import xlrd
import xlwt
import json
import requests
import threading
import logging


# sh = pandas.read_excel('BCA Test.xlsx')
# print(sh.columns)
# root = tk.Tk()
# new_book.save('new.xls')

def get_vrm_total():
    vrm_total = 0

    for row in range(4, sheet.nrows):
        if sheet.row_values(row)[2] == '':
            break
        else:
            print(sheet.row_values(row)[2])
            vrm_total += 1
    print(vrm_total)
    return vrm_total


def get_valuations(vrm_count):
    api_url_identity = 'https://staging.motorspecs.com/identity/lookup'

    headers_identity = {'Accept': 'application/vnd.identity.v2+json',
                        'Content-Type': 'application/vnd.identity.v2+json',
                        'Authorization': 'Bearer XXX',
                        'Connection': 'close'}

    currentRow = 4

    # vrm count was here

    for row in range(4, (4 + vrm_count)):
        temp_vrm = sheet.row_values(row)[2]
        current_vrm = temp_vrm.replace(" ", "")
        current_mileage = 0
        current_nat = sheet.row_values(row)[3]
        current_list_price = sheet.row_values(row)[4]

        body_identity = {'registration': current_vrm,
                         'currentMiles': current_mileage}

        response = requests.post(url=api_url_identity, json=body_identity, headers=headers_identity)

        response_dict = json.loads(response.content)
        print(response_dict)
        current_reg_date = response_dict['vehicle']['dvla']['regDate']

        if response.status_code != 201:
            notice = 'Error with VRM Lookup'
            new_sheet.write(currentRow, 16, notice)
        else:
            if current_nat == 'N/A':
                notice = 'No NAT code available'
                new_sheet.write(currentRow, 16, notice)
            else:
                # get_valuation(current_list_price, current_vrm, current_mileage, current_reg_date, current_nat,
                # currentRow, vrm_count)

                # temp_mileage = mileage

                api_url_valuation = 'https://staging.motorspecs.com/valuation-glass/by-nat-code/value'

                headers_valuation = {'Accept': 'application/vnd.valuation-glass.v2+json',
                                     'Content-Type': 'application/vnd.valuation-glass.v2+json',
                                     'Authorization': 'Bearer XXX',
                                     'Connection': 'close'}

                dates = ['2016-01-01', '2016-07-01', '2016-12-01', '2017-01-01', '2017-07-01', '2017-12-01',
                         '2018-01-01', '2018-07-01', '2018-12-01', '2019-01-01', '2019-07-01']

                current_date_col = 5
                got_new_price = False

                for date in dates:
                    write_value_row = row + vrm_count + 4
                    if date == '2016-01-01':
                        new_sheet.write(row, current_date_col, 100)
                        new_sheet2.write(row, current_date_col, 100)

                        # new_sheet.write(write_value_row, current_date_col, current_list_price)
                        # new_sheet2.write(write_value_row, current_date_col, current_list_price)
                    else:
                        body_valuation = {'country': 'uk',
                                          'natCode': current_nat,
                                          'firstRegDate': '2016-01-01',
                                          'valuationDate': date,
                                          'isCommercial': '0',
                                          'currentMiles': current_mileage}

                        response_valuation = requests.post(url=api_url_valuation, json=body_valuation,
                                                           headers=headers_valuation)

                        if response_valuation.status_code == 404:
                            print(
                                "No valuation available for " + current_vrm + " with ValDate " + date + " and RegDate " + current_reg_date)
                            new_sheet.write(row, current_date_col, 'N/A')
                            new_sheet2.write(row, current_date_col, 'N/A')

                            new_sheet.write(write_value_row, current_date_col, 'N/A')
                            new_sheet2.write(write_value_row, current_date_col, 'N/A')
                        else:
                            response_dict = json.loads(response_valuation.content)

                            print(response_dict)

                            glassValuations = response_dict['glassValuation'][0]
                            tradeValue = glassValuations['adjustedTradeValues']['trade']
                            retailValue = glassValuations['adjustedTradeValues']['retail']
                            newPrice = glassValuations['newPrice']

                            if got_new_price is not True:
                                # write new price in percentage area
                                new_sheet.write(row, 4, newPrice)
                                new_sheet2.write(row, 4, newPrice)
                                new_sheet.write(row, 5, 100)
                                new_sheet2.write(row, 5, 100)
                                # write new price in value area
                                new_sheet.write(write_value_row, 4, newPrice)
                                new_sheet2.write(write_value_row, 4, newPrice)
                                new_sheet.write(write_value_row, 5, newPrice)
                                new_sheet2.write(write_value_row, 5, newPrice)
                                got_new_price = True

                            trade_percentage = int((1 - ((newPrice - tradeValue) / newPrice)) * 100)
                            retail_percentage = int(
                                (1 - ((newPrice - retailValue) / newPrice)) * 100)

                            # write percentages
                            new_sheet.write(row, current_date_col, trade_percentage)
                            new_sheet2.write(row, current_date_col, retail_percentage)
                            # write values
                            new_sheet.write(write_value_row, current_date_col, tradeValue)
                            new_sheet2.write(write_value_row, current_date_col, retailValue)

                    current_mileage += 5000

                    current_date_col += 1

        currentRow += 1

    new_book.save(file_path + '.RESULTS.xls')
    new_filename = file_path + '.RESULTS.xls'
    return new_filename


def run_averages(filename, vrm_count):
    book_wvalues = xlrd.open_workbook(filename)
    sheet_wvalues = book_wvalues.sheet_by_index(0)
    sheet2_wvalues = book_wvalues.sheet_by_index(1)

    new_book_wvalues = copy(book_wvalues)
    new_sheet_wvalues = new_book_wvalues.get_sheet(0)
    new_sheet2_wvalues = new_book_wvalues.get_sheet(1)

    for column in range(4, 16):
        temp_sum_trade = 0
        temp_sum_retail = 0
        temp_sum_trade_value = 0
        temp_sum_retail_value = 0
        temp_count = 0
        for row in range(4, (4 + vrm_count)):
            if sheet_wvalues.row_values(row)[column] == 'N/A':
                continue
            elif sheet_wvalues.row_values(row)[column] == '':
                continue
            else:
                temp_sum_trade += sheet_wvalues.row_values(row)[column]
                temp_sum_retail += sheet2_wvalues.row_values(row)[column]
                temp_count += 1
            new_sheet_wvalues.write(4 + vrm_count, column, int(temp_sum_trade / temp_count))
            new_sheet2_wvalues.write(4 + vrm_count, column, int(temp_sum_retail / temp_count))
        for row in range((8 + vrm_count), (8 + (2 * vrm_count))):
            if sheet_wvalues.row_values(row)[column] == 'N/A':
                continue
            elif sheet_wvalues.row_values(row)[column] == '':
                continue
            else:
                temp_sum_trade_value += sheet_wvalues.row_values(row)[column]
                temp_sum_retail_value += sheet2_wvalues.row_values(row)[column]
            new_sheet_wvalues.write(8 + (2 * vrm_count), column, int(temp_sum_trade_value / temp_count))
            new_sheet2_wvalues.write(8 + (2 * vrm_count), column, int(temp_sum_retail_value / temp_count))

        row + vrm_count + 4

    new_book_wvalues.save(file_path + '.FINAL.xls')


def test_val():
    data = '''
    {"vehicleId":null,"registration":null,"currentMiles":null,"priceWhenNew":null,"valuationDate":"2018-01-01","glassValuation":[{"version":null,"modelId":"208583","modelQualifier":"001","qualifiedModelCode":208583001,"glassCode":null,"newPrice":27400,"averageMileage":22000,"basicValue":{"trade":12350,"retail":15600},"adjustedMileage":"20000","adjustedTradeValues":{"trade":12480,"retail":15730,"tradeHigh":12680,"tradeAverage":11930,"tradeLow":11620},"adjustedConsumerValues":{"partExExcellent":12330,"partExAverage":11120,"partExLow":10260,"retail":15730,"privateSale":13920,"retailTransacted":15130},"commercialVehicle":{"lowMileageTrade":0,"lowMileageRetail":0,"disposalTrade":null}}],"_links":{"self":{"href":"https:\/\/staging.motorspecs.com\/valuation-glass\/by-nat-code\/value"}}}
        '''
    response_dict2 = json.loads(data)

    # newPrice = response_dict2['glassValuation'][0]['newPrice']
    # basicValue = response_dict2['glassValuation'][0]['basicValue']
    tradeValues = response_dict2['glassValuation'][0]['adjustedTradeValues']

    # tradeValue = adjustedTradeValues['trade']
    # print(newPrice)
    # print(basicValue['trade'])
    # print(basicValue['retail'])
    print(tradeValues['trade'])
    print(tradeValues['retail'])


if __name__ == "__main__":
    root = tk.Tk()
    root.title('MotorCheck - Depreciation Checker')
    root.withdraw()
    file_path = tk.filedialog.askopenfilename()

    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)
    sheet2 = book.sheet_by_index(1)

    new_book = copy(book)
    new_sheet = new_book.get_sheet(0)
    new_sheet2 = new_book.get_sheet(1)
    new_sheet.write(0, 16, 'Notes')

    temp_vrm_count = get_vrm_total()
    temp_file = get_valuations(temp_vrm_count)
    run_averages(temp_file, temp_vrm_count)
