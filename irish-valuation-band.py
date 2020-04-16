from xlutils.copy import copy
import tkinter as tk
from tkinter import filedialog
import xlrd
import xlwt
import json
import pandas as pd
import requests
import requests_cache
import re
from bs4 import BeautifulSoup
import os
import shutil
from pathlib import Path
from datetime import datetime, time

requests_cache.install_cache('irish')


# Didn't finish this off...
def mergeFiles():
    temp_data_folder = input("Enter directory...\n")
    data_folder = Path(temp_data_folder)
    for filename in os.listdir(data_folder):
        if filename.endswith(".xls"):
            book = xlrd.open_workbook(data_folder / filename)
            sheet = book.sheet_by_index(0)
            new_book = copy(book)
            new_sheet = new_book.get_sheet(0)
            # currentRow = 0
            for row in range(sheet.nrows):
                new_sheet.write(row)
            print("Done: ", filename)
        else:
            print("No files to process")
    new_book.save(data_folder / 'FINAL.xls')


def getValuation():
    bands = {range(0, 1000): 'A',
             range(1001, 2000): 'B',
             range(2001, 3000): 'C',
             range(3001, 4000): 'D',
             range(4001, 5000): 'E',
             range(5001, 6000): 'F',
             range(6001, 7000): 'G',
             range(7001, 8000): 'H',
             range(8001, 9000): 'I',
             range(9001, 10000): 'J',
             range(10001, 11000): 'K',
             range(11001, 12000): 'L',
             range(12001, 13000): 'M',
             range(13001, 14000): 'N',
             range(14001, 15000): 'O',
             range(15001, 16000): 'P',
             range(16001, 17000): 'Q',
             range(17001, 18000): 'R',
             range(18001, 19000): 'S',
             range(19001, 20000): 'T',
             range(20001, 21000): 'U',
             range(21001, 22000): 'V',
             range(22001, 23000): 'W',
             range(23001, 24000): 'X',
             range(24001, 25000): 'Y',
             range(25001, 26000): 'Z',
             range(26001, 27000): 'A1',
             range(27001, 28000): 'B1',
             range(28001, 29000): 'C1',
             range(29001, 30000): 'D1',
             range(30001, 31000): 'E1',
             range(31001, 32000): 'F1',
             range(32001, 33000): 'G1',
             range(33001, 34000): 'H1',
             range(34001, 35000): 'I1',
             range(35001, 36000): 'J1',
             range(36001, 37000): 'K1',
             range(37001, 38000): 'L1',
             range(38001, 39000): 'M1',
             range(39001, 40000): 'N1',
             range(40001, 41000): 'O1',
             range(41001, 42000): 'P1',
             range(42001, 43000): 'Q1',
             range(43001, 44000): 'R1',
             range(44001, 45000): 'S1',
             range(45001, 46000): 'T1',
             range(46001, 47000): 'U1',
             range(47001, 48000): 'V1',
             range(48001, 49000): 'W1',
             range(49001, 50000): 'X1',
             range(50001, 51000): 'Y1',
             range(51001, 52000): 'Z1',
             range(52001, 53000): 'A2',
             range(53001, 54000): 'B2',
             range(54001, 55000): 'C2',
             range(55001, 56000): 'D2',
             range(56001, 57000): 'E2',
             range(57001, 58000): 'F2',
             range(58001, 59000): 'G2',
             range(59001, 60000): 'H2',
             range(60001, 61000): 'I2',
             range(61001, 62000): 'J2',
             range(62001, 63000): 'K2',
             range(63001, 64000): 'L2',
             range(64001, 65000): 'M2',
             range(65001, 66000): 'N2',
             range(66001, 67000): 'O2',
             range(67001, 68000): 'P2',
             range(68001, 69000): 'Q2',
             range(69001, 70000): 'R2',
             range(70001, 71000): 'S2',
             range(71001, 72000): 'T2',
             range(72001, 73000): 'U2',
             range(73001, 74000): 'V2',
             range(74001, 75000): 'W2',
             range(75001, 76000): 'X2',
             range(76001, 77000): 'Y2',
             range(77001, 78000): 'Z2',
             range(78001, 79000): 'A3',
             range(79001, 80000): 'B3',
             range(80001, 81000): 'C3',
             range(81001, 82000): 'D3',
             range(82001, 83000): 'E3',
             range(83001, 84000): 'F3',
             range(84001, 85000): 'G3',
             range(85001, 86000): 'H3',
             range(86001, 87000): 'I3',
             range(87001, 88000): 'J3',
             range(88001, 89000): 'K3',
             range(89001, 90000): 'L3',
             range(90001, 91000): 'M3',
             range(91001, 92000): 'N3',
             range(92001, 93000): 'O3',
             range(93001, 94000): 'P3',
             range(94001, 95000): 'Q3',
             range(95001, 96000): 'R3',
             range(96001, 97000): 'S3',
             range(97001, 98000): 'T3',
             range(98001, 99000): 'U3',
             range(99001, 100000): 'V3',
             range(100001, 200000): 'W3',
             range(200001, 300000): 'X3',
             range(300001, 400000): 'Y3',
             range(400001, 500000): 'Z3',
             range(500001, 1000000): 'A4',
             range(1000001, 999999999): 'B4'}
    now = datetime.now()
    now_time = now.time()

    temp_data_folder = input("Enter directory...\n")
    data_folder = Path(temp_data_folder)
    for filename in os.listdir(data_folder):
        if filename.endswith(".xlsx"):
            book = xlrd.open_workbook(data_folder / filename)
            sheet = book.sheet_by_index(0)
            new_book = copy(book)
            new_sheet = new_book.get_sheet(0)
            new_sheet.write(0, 10, 'Notes')
            currentRow = 0
            for row in range(sheet.nrows):
                if sheet.row_values(row)[0] == 'VRM' or sheet.row_values(row)[0] == 'MVRIS':
                    pass
                else:
                    vrm = sheet.row_values(row)[0]
                    api_url = 'https://api.motorcheck.ie/vehicle/reg/' + vrm + '/valuation?_end_user_ref=XXX&_username=XXX&_api_key=XXX'
                    response = requests.get(url=api_url).text
                    irish_response = BeautifulSoup(response, 'html.parser')
                    try:
                        value_market = irish_response.value_market.string
                        if value_market is None:
                            new_sheet.write(currentRow, 10, 'No valuation available')
                        else:
                            new_sheet.write(currentRow, 7, value_market)
                            for key in bands:
                                if int(value_market) in key:
                                    vehicle_band = bands.get(key)
                                    new_sheet.write(currentRow, 8, vehicle_band)
                    except AttributeError:
                        new_sheet.write(currentRow, 7, 0)
                    print('Done: ', vrm)
                currentRow = currentRow + 1

            # new_book.save(file_path + '.RESULTS.xls')
            temp_file = filename + '.RESULTS.xls'
            new_book.save(data_folder / 'output' / temp_file)
            shutil.move(data_folder / filename, data_folder / 'complete' / filename)
            print("Done: ", filename)
            if time(6, 00) >= now_time >= time(17, 00):
                return print("Halting script as not within the correct time")
        else:
            print("No files to process")
            exit

if __name__ == "__main__":
    # root = tk.Tk()
    # root.title('MotorCheck.ie Valuations')
    # root.withdraw()
    # file_path = tk.filedialog.askopenfilename()
    #
    # book = xlrd.open_workbook(file_path)
    # sheet = book.sheet_by_index(0)
    #
    # new_book = copy(book)
    # new_sheet = new_book.get_sheet(0)
    # new_sheet.write(0, 10, 'Notes')

    getValuation()
