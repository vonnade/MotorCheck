from xlutils.copy import copy
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext as ScrolledText
import xlrd
import xlwt
import json
import requests
import requests_cache
import re
import threading
import logging

# sh = pandas.read_excel('BCA Test.xlsx')
# print(sh.columns)

# root = tk.Tk()


# new_book.save('new.xls')


grant_type = 'client_credentials'
client_id = 'jbr-capital'
client_secret = '6bSRMMPa5n3GFev05nKm'

oauth_url = 'https://staging.motorspecs.com/oauth'

headers = {'Accept': 'application/json',
           'Content-Type': 'application/json',
           'Authorization': 'Bearer XXX'}


def fix_model():
    api_url = 'https://staging.motorspecs.com/identity-specs/lookup'

    headers_checkVRM = {'Accept': 'application/vnd.identity-specs.v2+json',
                        'Content-Type': 'application/vnd.identity-specs.v2+json',
                        'Authorization': 'Bearer XXX',
                        'Cache-Control': 'Cache-Control: max-age=300'}

    #    body = {'registration': 'yr12okh',
    #            'currentMiles': '37000'}
    currentRow = 0

    requests_cache.install_cache('check_vrm', backend='sqlite', expire_after=1800)

    for row in range(sheet.nrows):
        if sheet.row_values(row)[0] == 'VRM' or sheet.row_values(row)[0] == 'MVRIS':
            pass
        else:
            # current_VIN = sheet.row_values(row)[0]
            temp_vrm = sheet.row_values(row)[0]
            current_vrm = temp_vrm.replace(" ", "")
            current_mileage = 0

            body = {'registration': current_vrm,
                    'currentMiles': current_mileage}

            response = requests.post(url=api_url, json=body, headers=headers_checkVRM)

            # response_dict = json.loads(response.content)
            # print(response_dict)

            if response.status_code != 201:
                notice = 'Error with VRM Lookup'
                new_sheet.write(currentRow, 20, notice)
                # print(response_dict)
                # print(notice)
            else:
                response_dict = json.loads(response.content)
                # print(response_dict)

                # retrieved_VIN = response_dict['vehicle']['dvla']['vin']
                # retrieved_VIN = dvla_Records['vin']

                # if retrieved_VIN != current_VIN:
                #    notice = ('VIN does not match with DVLA records')
                #    new_sheet.write(currentRow, 3, notice)

                if response_dict['vehicle']['mvris']['model'] == '':
                    model = response_dict['vehicle']['dvla']['model']
                    new_sheet.write(currentRow, 2, model)
                else:
                    model = response_dict['vehicle']['mvris']['model']
                    new_sheet.write(currentRow, 2, model)

                # get_valuation(current_vrm, current_mileage, currentRow)

            print('Done: ', current_vrm, current_mileage)
        currentRow = currentRow + 1

    new_book.save(file_path + '.RESULTS.xls')


def check_vrm():
    # api_url = 'https://staging.motorspecs.com/provenance/check'
    api_url = 'https://staging.motorspecs.com/identity-specs/lookup'

    headers_checkVRM = {'Accept': 'application/vnd.identity-specs.v2+json',
                        'Content-Type': 'application/vnd.identity-specs.v2+json',
                        'Authorization': 'Bearer XXX',
                        'Cache-Control': 'Cache-Control: max-age=300'}

    #    body = {'registration': 'yr12okh',
    #            'currentMiles': '37000'}
    currentRow = 0

    requests_cache.install_cache('check_vrm', backend='sqlite', expire_after=1800)

    for row in range(sheet.nrows):
        if sheet.row_values(row)[0] == 'VRM' or sheet.row_values(row)[0] == 'MVRIS':
            pass
        else:
            # current_VIN = sheet.row_values(row)[0]
            temp_vrm = sheet.row_values(row)[0]
            current_vrm = temp_vrm.replace(" ", "")
            current_mileage = 0

            body = {'registration': current_vrm,
                    'currentMiles': current_mileage}

            response = requests.post(url=api_url, json=body, headers=headers_checkVRM)

            response_dict = json.loads(response.content)
            # print(response_dict)

            if response.status_code != 201:
                # notice = 'Error with VRM Lookup'
                notice = response_dict['detail']
                new_sheet.write(currentRow, 23, notice)
                # print(response_dict)
                # print(notice)
            else:
                # response_dict = json.loads(response.content)
                # print(response_dict)

                # retrieved_VIN = response_dict['vehicle']['dvla']['vin']
                # retrieved_VIN = dvla_Records['vin']

                # if retrieved_VIN != current_VIN:
                #    notice = ('VIN does not match with DVLA records')
                #    new_sheet.write(currentRow, 3, notice)

                if response_dict['vehicle']['dvla']['make'] == '':
                    make = response_dict['vehicle']['mvris']['make']
                    new_sheet.write(currentRow, 1, make)
                else:
                    make = response_dict['vehicle']['dvla']['make']
                    new_sheet.write(currentRow, 1, make)

                if response_dict['vehicle']['dvla']['model'] == '':
                    model = response_dict['vehicle']['mvris']['vehicleDesc']
                    new_sheet.write(currentRow, 2, model)
                else:
                    model = response_dict['vehicle']['dvla']['model']
                    new_sheet.write(currentRow, 2, model)

                if response_dict['vehicle']['dvla']['body'] == '':
                    bodyDesc = response_dict['vehicle']['mvris']['bodyDesc']
                    new_sheet.write(currentRow, 3, bodyDesc)
                else:
                    bodyDesc = response_dict['vehicle']['dvla']['body']
                    new_sheet.write(currentRow, 3, bodyDesc)

                gearboxType = response_dict['vehicle']['mvris']['gearboxType']
                new_sheet.write(currentRow, 4, gearboxType)

                if response_dict['vehicle']['dvla']['fuel'] == '':
                    fuel = response_dict['vehicle']['mvris']['fuel']
                    new_sheet.write(currentRow, 5, fuel)
                else:
                    fuel = response_dict['vehicle']['dvla']['fuel']
                    new_sheet.write(currentRow, 5, fuel)

                if response_dict['vehicle']['dvla']['cc'] == '':
                    cc = response_dict['vehicle']['mvris']['cc']
                    new_sheet.write(currentRow, 6, cc)
                else:
                    cc = response_dict['vehicle']['dvla']['cc']
                    new_sheet.write(currentRow, 6, cc)

                if response_dict['vehicle']['mvris']['engineSize'] == '':
                    engineSize = response_dict['vehicle']['dvla']['cc']
                    temp_engineSize = round((engineSize / 1000), 2)
                    new_sheet.write(currentRow, 7, temp_engineSize)
                else:
                    engineSize = response_dict['vehicle']['mvris']['engineSize']
                    new_sheet.write(currentRow, 7, engineSize)

                bhpCount = response_dict['vehicle']['mvris']['bhpCount']
                new_sheet.write(currentRow, 8, bhpCount)

                manufDate = response_dict['vehicle']['combined']['manufDate']
                try:
                    temp_manufDate = manufDate[:4]
                    new_sheet.write(row, 9, temp_manufDate)
                except:
                    new_sheet.write(currentRow, 9, manufDate)

                if response_dict['vehicle']['mvris']['regDate'] == '':
                    regDate = response_dict['vehicle']['dvla']['regDate']
                    new_sheet.write(currentRow, 10, regDate)
                else:
                    regDate = response_dict['vehicle']['mvris']['regDate']
                    new_sheet.write(currentRow, 10, regDate)

                mpg = response_dict['vehicle']['mvris']['combinedMpg']
                new_sheet.write(currentRow, 16, mpg)
                lkm = response_dict['vehicle']['mvris']['combinedLkm']
                new_sheet.write(currentRow, 17, lkm)
                dvla_co2 = response_dict['vehicle']['dvla']['co2']
                new_sheet.write(currentRow, 22, dvla_co2)
                mvris_co2 = response_dict['vehicle']['mvris']['vehicleCo2']
                new_sheet.write(currentRow, 15, mvris_co2)
                eu = response_dict['vehicle']['mvris']['euroStatus']
                new_sheet.write(currentRow, 11, eu)
                run_SpecCheck(current_vrm, currentRow, mvris_co2, mpg, lkm, eu)

                driveAxle = response_dict['vehicle']['mvris']['driveAxle']
                new_sheet.write(currentRow, 14, driveAxle)

                typeApprovalCategory = response_dict['vehicle']['mvris']['typeApprovalCategory']
                new_sheet.write(currentRow, 19, typeApprovalCategory)

                # get_valuation(current_vrm, current_mileage, currentRow)

            print('Done: ', current_vrm, current_mileage)
        currentRow = currentRow + 1

    new_book.save(file_path + '.RESULTS.xls')


def run_SpecCheck(vrm, row, co2, mpg, lkm, eu):
    # RUN SPEC CHECK
    api_url = 'https://staging.motorspecs.com/specs/standard'

    headers_speccheck = {'Accept': 'application/vnd.specs.v2+json',
                         'Content-Type': 'application/vnd.specs.v2+json',
                         'Authorization': 'Bearer XXX',
                         'Cache-Control': 'Cache-Control: max-age=300'}

    body = {'registration': vrm}

    requests_cache.install_cache('run_SpecCheck', backend='sqlite', expire_after=1800)

    response = requests.post(url=api_url, json=body, headers=headers_speccheck)

    response_dict = json.loads(response.content)

    if response.status_code != 201:
        notice = response_dict['detail']
        new_sheet.write(row, 23, notice)
        # print(notice)
    else:
        temp_dict = response_dict['standardSpecification']
        # print(temp_dict.get('Technical'))
        exact_item = 0
        for section in temp_dict:
            # print(section)
            if list(section.keys())[0] == 'Technical':
                # print('FOUND IT!!')
                break
            exact_item = exact_item + 1
        # print('Found at: ' + str(exact_item))

        tech_data = temp_dict[exact_item]
        # print(tech_data)

        emissions_items = ''
        combined_items = ''

        exact_item_2 = 0
        for tech_item in tech_data['Technical']:
            current_item = (list(tech_item.items()))
            # print(current_item[0])
            if str(current_item[0]) == "('id', 7601)":
                try:
                    emissions_items = current_item[4][1]
                except IndexError:
                    new_sheet.write(row, 23, 'No EU status available from Specs')
            if str(current_item[0]) == "('id', 42001)":
                try:
                    combined_items = current_item[4][1]
                    break
                except IndexError:
                    new_sheet.write(row, 24, 'No fuel consumption data available from Specs')
            exact_item_2 = exact_item_2 + 1

        try:
            if emissions_items[0]['id'] == 7602:
                try:
                    eu_rating = re.sub('[^0-9\.]', '', emissions_items[0]['value'])
                    new_sheet.write(row, 21, eu_rating)
                except IndexError:
                    new_sheet.write(row, 23, 'No EU status available from Specs')
                # if co2 is None:
                try:
                    if emissions_items[1]['id'] == 7603:
                        co2_value = emissions_items[1]['value']
                        new_sheet.write(row, 20, co2_value)
                    else:
                        new_sheet.write(row, 23, 'No CO2 value available from Specs')
                except IndexError:
                    new_sheet.write(row, 23, 'No CO2 value available from Specs')
            elif emissions_items[0]['id'] == 7603:
                try:
                    co2_value = emissions_items[0]['value']
                    new_sheet.write(row, 20, co2_value)
                except IndexError:
                    new_sheet.write(row, 23, 'No CO2 value available from Specs')
            else:
                new_sheet.write(row, 23, 'No EU status available from Specs')
        except IndexError:
            new_sheet.write(row, 23, 'No EU status available from Specs')

        # try:
        #     if emissions_items[0]['id'] == 7603:
        #         try:
        #             co2_value = emissions_items[0]['value']
        #             new_sheet.write(row, 15, co2_value)
        #         except IndexError:
        #             new_sheet.write(row, 23, 'No CO2 value available from DVLA or Specs')
        #     elif emissions_items[1]['id'] == 7603:
        #         try:
        #             co2_value = emissions_items[1]['value']
        #             new_sheet.write(row, 15, co2_value)
        #         except IndexError:
        #             new_sheet.write(row, 23, 'No CO2 value available from DVLA or Specs')
        # except IndexError:
        #     new_sheet.write(row, 23, 'No CO2 value available from DVLA or Specs')

        # try:
        #     for item in combined_items:
        #         # print(item['id'])
        #         if lkm is None:
        #             if item['id'] == 42005:
        #                 spec_lkm = item['value']
        #                 new_sheet.write(row, 17, spec_lkm)
        #         else:
        #             new_sheet.write(row, 17, lkm)
        #
        #         if mpg is None:
        #             if item['id'] == 142008:
        #                 spec_mpg = item['value']
        #                 new_sheet.write(row, 16, spec_mpg)
        #         else:
        #             new_sheet.write(row, 16, mpg)
        # except:
        #     return


def run_Identity(vrm):
    api_url_2 = 'https://staging.motorspecs.com/identity-specs/lookup'

    headers_2 = {'Accept': 'application/vnd.identity-specs.v2+json',
                 'Content-Type': 'application/vnd.identity-specs.v2+json',
                 'Authorization': 'Bearer XXX'}

    body_2 = {'registration': vrm,
              'currentMiles': 0}

    response = requests.post(url=api_url_2, json=body_2, headers=headers_2)

    return


def get_euroStatus():
    print('called euro def')
    currentRow = 0

    for row in range(sheet.nrows):
        if sheet.row_values(row)[0] == 'VRM' or sheet.row_values(row)[0] == 'MVRIS':
            pass
        else:
            temp_vrm = sheet.row_values(row)[0]
            current_vrm = temp_vrm.replace(" ", "")
            # temp_mileage = sheet.row_values(row)[1]
            # current_mileage = int(temp_mileage)
            current_mileage = 0
            print(current_vrm)

            if sheet.row_values(row)[11] == '':
                print('No EU status for: ', current_vrm)

                run_Identity(current_vrm)

                # RUN SPEC CHECK
                api_url = 'https://staging.motorspecs.com/specs/standard'

                headers = {'Accept': 'application/vnd.specs.v2+json',
                           'Content-Type': 'application/vnd.specs.v2+json',
                           'Authorization': 'Bearer XXX'}

                body = {'registration': current_vrm}

                response = requests.post(url=api_url, json=body, headers=headers)

                response_dict = json.loads(response.content)

                if response.status_code != 201:
                    notice = ("No specs available")
                    new_sheet.write(row, 20, notice)
                    # print(notice)
                else:
                    temp_dict = response_dict['standardSpecification']
                    # print(temp_dict.get('Technical'))
                    exact_item = 0
                    for section in temp_dict:
                        # print(section)
                        if list(section.keys())[0] == 'Technical':
                            # print('FOUND IT!!')
                            break
                        exact_item = exact_item + 1
                    # print('Found at: ' + str(exact_item))

                    tech_data = temp_dict[exact_item]
                    # print(tech_data)

                    exact_item_2 = 0
                    for item in tech_data['Technical']:
                        current_item = (list(item.items()))
                        print(current_item[0])
                        if str(current_item[0]) == "('id', 7601)":
                            emissions_items = current_item[4][1]
                            break
                        exact_item_2 = exact_item_2 + 1

                    try:
                        eu_rating = emissions_items[0]['value']
                        new_sheet.write(row, 11, eu_rating)
                    except IndexError:
                        new_sheet.write(row, 20, 'No EU rating available from Spec Check')

                    try:
                        co2_value = emissions_items[1]['value']
                        new_sheet.write(row, 15, co2_value)
                    except IndexError:
                        new_sheet.write(row, 20, 'No CO2 value available from Spec Check')

                    print('Done: ', current_vrm)
            else:
                pass

        currentRow = currentRow + 1

    new_book.save(file_path + '.SPECRESULTS.xls')


def dict_Test():
    dict = {
        "id": 1,
        "registration": "AD52NHO",
        "vehicleId": "F85FT95IB1tbhQ3Xx0Mz9OGIy4H2IZBVOXdhLo8KE78=",
        "priceData": {
            "priceDate": "2002-11-07",
            "priceDateFormated": "11/2002",
            "priceDateMonth": "11",
            "priceDateYear": 2002,
            "price": 12075,
            "msrp": 12075,
            "priceOtr": 12245,
            "deliveryCharge": 170,
            "currency": "GBP",
            "currencySymbol": "&pound;"
        },
        "topFeatures": [
            {
                "id": 4001,
                "name": "Remote boot/hatch/rear door release",
                "description": "Electric remote boot/hatch/rear door release",
                "value": "standard"
            },
            {
                "id": 4301,
                "name": "Immobiliser",
                "description": "Immobiliser",
                "value": "standard"
            },
            {
                "id": 13001,
                "name": "Front fog lights",
                "description": "Front fog lights",
                "value": "standard"
            },
            {
                "id": 14801,
                "name": "Central door locking",
                "description": "Remote central door locking includes dead bolt",
                "value": "standard"
            },
            {
                "id": 16301,
                "name": "Front airbag",
                "description": "Driver and passenger front airbag intelligent",
                "value": "standard"
            },
            {
                "id": 18501,
                "name": "Power steering",
                "description": "Power steering",
                "value": "standard"
            },
            {
                "id": 19601,
                "name": "Cup holders",
                "description": "Cup holders for front seats fixed",
                "value": "standard"
            },
            {
                "id": 23301,
                "name": "Electric windows",
                "description": "Front electric windows with one one-touch",
                "value": "standard"
            }
        ],
        "standardSpecification": [
            {
                "Interior": [
                    {
                        "id": 701,
                        "name": "Seating",
                        "description": "Seating: five seats",
                        "value": "standard",
                        "items": [
                            {
                                "id": 702,
                                "name": "Seating capacity",
                                "value": 5
                            }
                        ]
                    },
                    {
                        "id": 1101,
                        "name": "Speakers",
                        "description": "Four speakers",
                        "value": "standard",
                        "items": [
                            {
                                "id": 1102,
                                "name": "number of",
                                "value": 4
                            }
                        ]
                    },
                    {
                        "id": 1301,
                        "name": "Audio player",
                        "description": "RDS ARI/EON audio player with AM/FM and CD player",
                        "value": "standard",
                        "items": [
                            {
                                "id": 1302,
                                "name": "radio",
                                "value": "AM/FM"
                            },
                            {
                                "id": 1304,
                                "name": "in-dash CD",
                                "value": "yes"
                            },
                            {
                                "id": 1309,
                                "name": "RDS",
                                "value": "yes"
                            },
                            {
                                "id": 1310,
                                "name": "ARI/EON",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 4001,
                        "name": "Remote boot/hatch/rear door release",
                        "description": "Electric remote boot/hatch/rear door release",
                        "value": "standard",
                        "items": [
                            {
                                "id": 4002,
                                "name": "operation",
                                "value": "electric"
                            }
                        ]
                    },
                    {
                        "id": 4701,
                        "name": "Cigar lighter",
                        "description": "Front seats cigar lighter",
                        "value": "standard"
                    },
                    {
                        "id": 4901,
                        "name": "Courtesy lights",
                        "description": "Delayed/fade courtesy lights",
                        "value": "standard",
                        "items": [
                            {
                                "id": 4903,
                                "name": "delayed/fade",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 5501,
                        "name": "Vanity mirror",
                        "description": "Driver and passenger vanity mirror",
                        "value": "standard"
                    },
                    {
                        "id": 9501,
                        "name": "Tachometer",
                        "description": "Tachometer",
                        "value": "standard"
                    },
                    {
                        "id": 10701,
                        "name": "Low fuel level warning",
                        "description": "Low fuel level warning",
                        "value": "standard"
                    },
                    {
                        "id": 11401,
                        "name": "Headlight on warning sound",
                        "description": "Headlight on warning sound",
                        "value": "standard"
                    },
                    {
                        "id": 11901,
                        "name": "Luxury trim",
                        "description": "Luxury trim alloy on gearknob, titanium on doors and titanium on dashboard",
                        "value": "standard",
                        "items": [
                            {
                                "id": 11903,
                                "name": "on gearknob",
                                "value": "alloy"
                            },
                            {
                                "id": 11906,
                                "name": "on doors",
                                "value": "titanium"
                            },
                            {
                                "id": 11907,
                                "name": "on dashboard",
                                "value": "titanium"
                            }
                        ]
                    },
                    {
                        "id": 14701,
                        "name": "Load restraint",
                        "description": "Load restraint hooks",
                        "value": "standard",
                        "items": [
                            {
                                "id": 14702,
                                "name": "type",
                                "value": "hooks"
                            }
                        ]
                    },
                    {
                        "id": 14801,
                        "name": "Central door locking",
                        "description": "Remote central door locking includes dead bolt",
                        "value": "standard",
                        "items": [
                            {
                                "id": 14802,
                                "name": "operation",
                                "value": "remote"
                            },
                            {
                                "id": 14808,
                                "name": "includes dead bolt",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 17401,
                        "name": "Seat upholstery",
                        "description": "Cloth seat upholstery with additional cloth",
                        "value": "standard",
                        "items": [
                            {
                                "id": 17402,
                                "name": "main seat material",
                                "value": "cloth"
                            },
                            {
                                "id": 17403,
                                "name": "additional seat material",
                                "value": "cloth"
                            }
                        ]
                    },
                    {
                        "id": 17801,
                        "name": "Front seat",
                        "description": "Sports driver seat with height adjustment , sports passenger seat",
                        "value": "standard",
                        "items": [
                            {
                                "id": 17803,
                                "name": "type",
                                "location": "Driver",
                                "value": "sports"
                            },
                            {
                                "id": 17803,
                                "name": "type",
                                "value": "sports"
                            },
                            {
                                "id": 17807,
                                "name": "height adjustment",
                                "location": "Driver",
                                "value": "yes"
                            },
                            {
                                "id": 17813,
                                "name": "number of electrical adjustments",
                                "location": "Driver",
                                "value": "-"
                            },
                            {
                                "id": 17813,
                                "name": "number of electrical adjustments",
                                "value": "-"
                            }
                        ]
                    },
                    {
                        "id": 17901,
                        "name": "Rear seats",
                        "description": "Three asymmetrical split bench front facing split squab rear seats",
                        "value": "standard",
                        "items": [
                            {
                                "id": 17903,
                                "name": "type",
                                "location": "Front",
                                "value": "split bench"
                            },
                            {
                                "id": 17912,
                                "name": "folding",
                                "location": "Front",
                                "value": "asymmetrical"
                            },
                            {
                                "id": 17913,
                                "name": "squab flip-up",
                                "location": "Front",
                                "value": "split"
                            },
                            {
                                "id": 17915,
                                "name": "seating capacity",
                                "location": "Front",
                                "value": 3
                            }
                        ]
                    },
                    {
                        "id": 18401,
                        "name": "Steering wheel",
                        "description": "Leather covered steering wheel with tilt adjustment and telescopic adjustment",
                        "value": "standard",
                        "items": [
                            {
                                "id": 18402,
                                "name": "type",
                                "value": "leather covered"
                            },
                            {
                                "id": 18406,
                                "name": "height adjustment",
                                "value": "yes"
                            },
                            {
                                "id": 18407,
                                "name": "telescopic adjustment",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 19301,
                        "name": "Door pockets/bins",
                        "description": "Door pockets/bins for driver seat and passenger seat",
                        "value": "standard"
                    },
                    {
                        "id": 19401,
                        "name": "Seat back storage",
                        "description": "Front seat back storage",
                        "value": "standard"
                    },
                    {
                        "id": 19601,
                        "name": "Cup holders",
                        "description": "Cup holders for front seats fixed",
                        "value": "standard",
                        "items": [
                            {
                                "id": 19603,
                                "name": "type",
                                "location": "Front",
                                "value": "fixed"
                            }
                        ]
                    },
                    {
                        "id": 20801,
                        "name": "Ventilation system",
                        "description": "Ventilation system with air filter",
                        "value": "standard",
                        "items": [
                            {
                                "id": 20809,
                                "name": "air filter",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 21501,
                        "name": "Rear view mirror",
                        "description": "Rear view mirror",
                        "value": "standard"
                    },
                    {
                        "id": 23301,
                        "name": "Electric windows",
                        "description": "Front electric windows with one one-touch",
                        "value": "standard",
                        "items": [
                            {
                                "id": 23307,
                                "name": "number of one touch",
                                "location": "Front",
                                "value": 1
                            }
                        ]
                    },
                    {
                        "id": 26601,
                        "name": "Console",
                        "description": "Partial dashboard console with open storage box",
                        "value": "standard",
                        "items": [
                            {
                                "id": 26603,
                                "name": "type",
                                "location": "Driver",
                                "value": "partial"
                            },
                            {
                                "id": 26604,
                                "name": "storage",
                                "location": "Driver",
                                "value": "open"
                            }
                        ]
                    }
                ]
            },
            {
                "Exterior": [
                    {
                        "id": 1201,
                        "name": "Aerial",
                        "description": "Roof aerial",
                        "value": "standard",
                        "items": [
                            {
                                "id": 1202,
                                "name": "type",
                                "value": "roof"
                            }
                        ]
                    },
                    {
                        "id": 1501,
                        "name": "Coefficient of drag",
                        "description": "Coefficient of drag: 0.32",
                        "value": "standard",
                        "items": [
                            {
                                "id": 1502,
                                "name": "coefficient of drag",
                                "value": 0.32
                            }
                        ]
                    },
                    {
                        "id": 1801,
                        "name": "Side rubbing strip",
                        "description": "Side rubbing strip",
                        "value": "standard"
                    },
                    {
                        "id": 3301,
                        "name": "Bumpers",
                        "description": "Part-painted front and rear bumpers",
                        "value": "standard",
                        "items": [
                            {
                                "id": 3305,
                                "name": "colour",
                                "location": "Front",
                                "value": "painted"
                            },
                            {
                                "id": 3305,
                                "name": "colour",
                                "location": "Rear",
                                "value": "painted"
                            }
                        ]
                    },
                    {
                        "id": 13001,
                        "name": "Front fog lights",
                        "description": "Front fog lights",
                        "value": "standard"
                    },
                    {
                        "id": 13601,
                        "name": "High mount brake light",
                        "description": "High mount brake light",
                        "value": "standard"
                    },
                    {
                        "id": 14101,
                        "name": "Tyres",
                        "description": "Front and rear tyres with 195 mm tyre width, 60% tyre profile and V tyre rating",
                        "value": "standard",
                        "items": [
                            {
                                "id": 14103,
                                "name": "tyre width",
                                "location": "Front",
                                "value": 195
                            },
                            {
                                "id": 14103,
                                "name": "tyre width",
                                "location": "Rear",
                                "value": 195
                            },
                            {
                                "id": 14104,
                                "name": "tyre profile",
                                "location": "Front",
                                "value": 60
                            },
                            {
                                "id": 14104,
                                "name": "tyre profile",
                                "location": "Rear",
                                "value": 60
                            },
                            {
                                "id": 14105,
                                "name": "tyre speed rating",
                                "location": "Front",
                                "value": "V"
                            },
                            {
                                "id": 14105,
                                "name": "tyre speed rating",
                                "location": "Rear",
                                "value": "V"
                            }
                        ]
                    },
                    {
                        "id": 15201,
                        "name": "Paint",
                        "description": "Gloss paint",
                        "value": "standard",
                        "items": [
                            {
                                "id": 15202,
                                "name": "type",
                                "value": "gloss"
                            }
                        ]
                    },
                    {
                        "id": 21601,
                        "name": "Door mirrors",
                        "description": "Driver and passenger internally adjustable partial-painted door mirrors",
                        "value": "standard",
                        "items": [
                            {
                                "id": 21603,
                                "name": "type",
                                "location": "Driver",
                                "value": "internally adjustable"
                            },
                            {
                                "id": 21603,
                                "name": "type",
                                "value": "internally adjustable"
                            },
                            {
                                "id": 21607,
                                "name": "colour",
                                "location": "Driver",
                                "value": "painted"
                            },
                            {
                                "id": 21607,
                                "name": "colour",
                                "value": "painted"
                            }
                        ]
                    },
                    {
                        "id": 22301,
                        "name": "Rear windscreen",
                        "description": "Rear windscreen with intermittent",
                        "value": "standard",
                        "items": [
                            {
                                "id": 22306,
                                "name": "wipers",
                                "value": "intermittent"
                            }
                        ]
                    },
                    {
                        "id": 22801,
                        "name": "Windscreen wipers",
                        "description": "Windscreen wipers",
                        "value": "standard"
                    },
                    {
                        "id": 24401,
                        "name": "Wheels",
                        "description": "Front and rear alloy wheels with 15 inch rim diam, 6 inch rim width and partial wheel covers",
                        "value": "standard",
                        "items": [
                            {
                                "id": 24404,
                                "name": "rim type",
                                "location": "Front",
                                "value": "alloy"
                            },
                            {
                                "id": 24404,
                                "name": "rim type",
                                "location": "Rear",
                                "value": "alloy"
                            },
                            {
                                "id": 24405,
                                "name": "rim diameter (in)",
                                "location": "Front",
                                "value": 15
                            },
                            {
                                "id": 24405,
                                "name": "rim diameter (in)",
                                "location": "Rear",
                                "value": 15
                            },
                            {
                                "id": 24406,
                                "name": "rim width (in)",
                                "location": "Front",
                                "value": 6
                            },
                            {
                                "id": 24406,
                                "name": "rim width (in)",
                                "location": "Rear",
                                "value": 6
                            },
                            {
                                "id": 24414,
                                "name": "wheel covers",
                                "location": "Front",
                                "value": "partial"
                            },
                            {
                                "id": 24414,
                                "name": "wheel covers",
                                "location": "Rear",
                                "value": "partial"
                            }
                        ]
                    },
                    {
                        "id": 24501,
                        "name": "Spare wheel",
                        "description": "Space saver steel rim internal spare wheel",
                        "value": "standard",
                        "items": [
                            {
                                "id": 24502,
                                "name": "type",
                                "value": "space saver"
                            },
                            {
                                "id": 24503,
                                "name": "rim type",
                                "value": "steel"
                            }
                        ]
                    },
                    {
                        "id": 24601,
                        "name": "Non-corrosive body",
                        "description": "Full galvanised non-corrosive body",
                        "value": "standard",
                        "items": [
                            {
                                "id": 24602,
                                "name": "type",
                                "value": "galvanised"
                            },
                            {
                                "id": 24603,
                                "location": "full"
                            }
                        ]
                    }
                ]
            },
            {
                "Dimensions": [
                    {
                        "id": 5801,
                        "name": "External dimensions",
                        "description": "External dimensions: overall length (mm): 4,152, overall length (inches): 163.5, overall width (mm): 1,702, overall width (inches): 67, overall height (mm): 1,430, overall height (inches): 56.3, wheelbase (mm): 2,615, wheelbase (inches): 103, front track (mm): 1,494, front track (inches): 58.8, rear track (mm): 1,487, rear track (inches): 58.5 and wall to wall turning circle (mm): 10,900",
                        "value": "standard",
                        "items": [
                            {
                                "id": 5802,
                                "name": "overall length (mm)",
                                "value": 4152
                            },
                            {
                                "id": 5803,
                                "name": "overall width (mm)",
                                "value": 1702
                            },
                            {
                                "id": 5804,
                                "name": "overall height (mm)",
                                "value": 1430
                            },
                            {
                                "id": 5806,
                                "name": "wheelbase (mm)",
                                "value": 2615
                            },
                            {
                                "id": 5807,
                                "name": "front track (mm)",
                                "value": 1494
                            },
                            {
                                "id": 5808,
                                "name": "rear track (mm)",
                                "value": 1487
                            },
                            {
                                "id": 5810,
                                "name": "wall to wall turning circle (mm)",
                                "value": 10900
                            },
                            {
                                "id": 105802,
                                "name": "overall length (in)",
                                "value": 163.5
                            },
                            {
                                "id": 105803,
                                "name": "overall width (in)",
                                "value": 67
                            },
                            {
                                "id": 105804,
                                "name": "overall height (in)",
                                "value": 56.3
                            },
                            {
                                "id": 105806,
                                "name": "wheelbase (in)",
                                "value": 103
                            },
                            {
                                "id": 105807,
                                "name": "front track (in)",
                                "value": 58.8
                            },
                            {
                                "id": 105808,
                                "name": "rear track (in)",
                                "value": 58.5
                            },
                            {
                                "id": 105810,
                                "name": "wall to wall turning circle (ft)",
                                "value": 35.8
                            }
                        ]
                    },
                    {
                        "id": 5901,
                        "name": "Internal dimensions",
                        "description": "Internal dimensions: front headroom (mm): 995, front headroom (inches): 39.2, rear headroom (mm): 982, rear headroom (inches): 38.7, front leg room (mm): 1,095, front leg room (inches): 43.1, rear leg room (mm): 882, rear leg room (inches): 34.7, front shoulder room (mm): 1,358, front shoulder room (inches): 53.5, rear shoulder room (mm): 1,358 and rear shoulder room (inches): 53.5",
                        "value": "standard",
                        "items": [
                            {
                                "id": 5902,
                                "name": "headroom front (mm)",
                                "value": 995
                            },
                            {
                                "id": 5903,
                                "name": "headroom rear (mm)",
                                "value": 982
                            },
                            {
                                "id": 5906,
                                "name": "leg room front (mm)",
                                "value": 1095
                            },
                            {
                                "id": 5907,
                                "name": "leg room rear (mm)",
                                "value": 882
                            },
                            {
                                "id": 5908,
                                "name": "shoulder room front (mm)",
                                "value": 1358
                            },
                            {
                                "id": 5909,
                                "name": "shoulder room rear (mm)",
                                "value": 1358
                            },
                            {
                                "id": 105902,
                                "name": "headroom front (in)",
                                "value": 39.2
                            },
                            {
                                "id": 105903,
                                "name": "headroom rear (in)",
                                "value": 38.7
                            },
                            {
                                "id": 105906,
                                "name": "leg room front (in)",
                                "value": 43.1
                            },
                            {
                                "id": 105907,
                                "name": "leg room rear (in)",
                                "value": 34.7
                            },
                            {
                                "id": 105908,
                                "name": "shoulder room front (in)",
                                "value": 53.5
                            },
                            {
                                "id": 105909,
                                "name": "shoulder room rear (in)",
                                "value": 53.5
                            }
                        ]
                    },
                    {
                        "id": 6001,
                        "name": "Load compartment capacity",
                        "description": "Load compartment capacity: rear seat up; to lower window (litres): 350, rear seat up; to lower window (cu ft): 12.4, rear seat down (litres): 1,205 and rear seat down (cu ft): 42.6",
                        "value": "standard",
                        "items": [
                            {
                                "id": 6002,
                                "name": "rear seat up to lower window (l)",
                                "value": 350
                            },
                            {
                                "id": 6004,
                                "name": "rear seat down to roof (l)",
                                "value": 1205
                            },
                            {
                                "id": 106002,
                                "name": "rear seat up to lower window (cu ft)",
                                "value": 12.4
                            },
                            {
                                "id": 106004,
                                "name": "rear seat down to roof (cu ft)",
                                "value": 42.6
                            }
                        ]
                    },
                    {
                        "id": 8901,
                        "name": "Fuel tank",
                        "description": "55 litre 14.5 gallon main unleaded fuel tank",
                        "value": "standard",
                        "items": [
                            {
                                "id": 8903,
                                "name": "capacity (l)",
                                "value": 55
                            },
                            {
                                "id": 8904,
                                "name": "fuel type",
                                "value": "unleaded"
                            },
                            {
                                "id": 108903,
                                "name": "capacity (gal)",
                                "value": 14.5
                            }
                        ]
                    },
                    {
                        "id": 24101,
                        "name": "Weights",
                        "description": "Weights: gross vehicle weight rating (kg): 1,590, gross vehicle weight rating (lbs): 3,505, kerb weight (kg): 1,086, kerb weight (lbs): 2,394, gross trailer weight braked (kg): 1,200, gross trailer weight braked (lbs): 2,646 and kerb weight includes driver: kerb weight includes driver",
                        "value": "standard",
                        "items": [
                            {
                                "id": 24102,
                                "name": "gross vehicle weight (kg)",
                                "value": 1590
                            },
                            {
                                "id": 24103,
                                "name": "published kerb weight (kg)",
                                "value": 1086
                            },
                            {
                                "id": 24105,
                                "name": "gross trailer weight braked (kg)",
                                "value": 1200
                            },
                            {
                                "id": 24112,
                                "name": "kerb weight includes driver",
                                "value": "yes"
                            },
                            {
                                "id": 124102,
                                "name": "gross vehicle weight (lbs)",
                                "value": 3505
                            },
                            {
                                "id": 124103,
                                "name": "published kerb weight (lbs)",
                                "value": 2394
                            },
                            {
                                "id": 124105,
                                "name": "gross trailer weight braked (lbs)",
                                "value": 2646
                            }
                        ]
                    }
                ]
            },
            {
                "Safety": [
                    {
                        "id": 3101,
                        "name": "Disc brakes",
                        "description": "Two disc brakes including two ventilated discs",
                        "value": "standard",
                        "items": [
                            {
                                "id": 3102,
                                "name": "number of",
                                "value": 2
                            },
                            {
                                "id": 3103,
                                "name": "number of ventilated discs",
                                "value": 2
                            }
                        ]
                    },
                    {
                        "id": 12501,
                        "name": "Headlights",
                        "description": "Twin complex surface lens halogen bulb headlights",
                        "value": "standard",
                        "items": [
                            {
                                "id": 12502,
                                "name": "lens type",
                                "value": "complex surface"
                            },
                            {
                                "id": 12503,
                                "name": "bulb type (low beam)",
                                "value": "halogen"
                            },
                            {
                                "id": 12504,
                                "name": "configuration",
                                "value": "twin"
                            }
                        ]
                    },
                    {
                        "id": 12601,
                        "name": "Headlight control",
                        "description": "Headlight control with internal height adjustment",
                        "value": "standard",
                        "items": [
                            {
                                "id": 12602,
                                "name": "internal height adjustment",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 16301,
                        "name": "Front airbag",
                        "description": "Driver and passenger front airbag intelligent",
                        "value": "standard",
                        "items": [
                            {
                                "id": 16306,
                                "name": "intelligent",
                                "location": "Driver",
                                "value": "yes"
                            },
                            {
                                "id": 16306,
                                "name": "intelligent",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 16501,
                        "name": "Head restraints",
                        "description": "Two height adjustable head restraints on front seats , three height adjustable head restraints on rear seats",
                        "value": "standard",
                        "items": [
                            {
                                "id": 16504,
                                "name": "height adjustable",
                                "location": "Front",
                                "value": "yes"
                            },
                            {
                                "id": 16504,
                                "name": "height adjustable",
                                "location": "Rear",
                                "value": "yes"
                            },
                            {
                                "id": 16508,
                                "name": "number",
                                "location": "Front",
                                "value": 2
                            },
                            {
                                "id": 16508,
                                "name": "number",
                                "location": "Rear",
                                "value": 3
                            }
                        ]
                    },
                    {
                        "id": 16701,
                        "name": "Front seat belts",
                        "description": "Height adjustable 3-point reel front seat belts on driver seat and passenger seat with pre-tensioners",
                        "value": "standard",
                        "items": [
                            {
                                "id": 16703,
                                "name": "type",
                                "location": "Driver",
                                "value": "3-point"
                            },
                            {
                                "id": 16703,
                                "name": "type",
                                "value": "3-point"
                            },
                            {
                                "id": 16704,
                                "name": "operation",
                                "location": "Driver",
                                "value": "reel"
                            },
                            {
                                "id": 16704,
                                "name": "operation",
                                "value": "reel"
                            },
                            {
                                "id": 16706,
                                "name": "pre-tensioners",
                                "location": "Driver",
                                "value": "yes"
                            },
                            {
                                "id": 16706,
                                "name": "pre-tensioners",
                                "value": "yes"
                            },
                            {
                                "id": 16707,
                                "name": "height adjustable",
                                "location": "Driver",
                                "value": "yes"
                            },
                            {
                                "id": 16707,
                                "name": "height adjustable",
                                "value": "yes"
                            }
                        ]
                    },
                    {
                        "id": 16801,
                        "name": "Rear seat belts",
                        "description": "3-point reel rear seat belts on driver side, passenger side and centre side",
                        "value": "standard",
                        "items": [
                            {
                                "id": 16803,
                                "name": "type",
                                "location": "Driver",
                                "value": "3-point"
                            },
                            {
                                "id": 16803,
                                "name": "type",
                                "value": "3-point"
                            },
                            {
                                "id": 16803,
                                "name": "type",
                                "value": "3-point"
                            },
                            {
                                "id": 16804,
                                "name": "operation",
                                "location": "Driver",
                                "value": "reel"
                            },
                            {
                                "id": 16804,
                                "name": "operation",
                                "value": "reel"
                            },
                            {
                                "id": 16804,
                                "name": "operation",
                                "value": "reel"
                            }
                        ]
                    }
                ]
            },
            {
                "Security": [
                    {
                        "id": 4301,
                        "name": "Immobiliser",
                        "description": "Immobiliser",
                        "value": "standard"
                    },
                    {
                        "id": 26101,
                        "name": "Audio anti-theft protection",
                        "description": "Audio anti-theft protection: code, integrated into fascia and detachable panel",
                        "value": "standard"
                    }
                ]
            },
            {
                "Technical": [
                    {
                        "id": 6501,
                        "name": "Drive",
                        "description": "Front-wheel drive",
                        "value": "standard",
                        "items": [
                            {
                                "id": 6502,
                                "name": "Driven wheels",
                                "value": "front"
                            }
                        ]
                    },
                    {
                        "id": 7401,
                        "name": "Engine",
                        "description": "1,596 cc 1.6 litres in-line 4 transverse engine with 79 mm bore, 81.4 mm stroke, 11 compression ratio, double overhead cam and four valves per cylinder ZETEC",
                        "value": "standard",
                        "items": [
                            {
                                "id": 7402,
                                "name": "cc",
                                "value": 1596
                            },
                            {
                                "id": 7403,
                                "name": "Litres",
                                "value": 1.6
                            },
                            {
                                "id": 7404,
                                "name": "bore",
                                "value": 79
                            },
                            {
                                "id": 7405,
                                "name": "stroke",
                                "value": 81.4
                            },
                            {
                                "id": 7406,
                                "name": "compression ratio",
                                "value": 11
                            },
                            {
                                "id": 7407,
                                "name": "number of cylinders",
                                "value": 4
                            },
                            {
                                "id": 7408,
                                "name": "configuration",
                                "value": "in-line"
                            },
                            {
                                "id": 7411,
                                "name": "orientation",
                                "value": "transverse"
                            },
                            {
                                "id": 7414,
                                "name": "valve gear type",
                                "value": "double overhead cam"
                            },
                            {
                                "id": 7417,
                                "name": "number of valves per cylinder",
                                "value": 4
                            },
                            {
                                "id": 7420,
                                "name": "engine code",
                                "value": "ZETEC"
                            }
                        ]
                    },
                    {
                        "id": 7601,
                        "name": "Emission control level",
                        "description": "Emission control level EU3 - carbon dioxide level (g/km): 169",
                        "value": "standard",
                        "items": [
                            {
                                "id": 7602,
                                "name": "standard met",
                                "value": "EU3"
                            },
                            {
                                "id": 7603,
                                "name": "CO2 level - g/km combined",
                                "value": 169
                            }
                        ]
                    },
                    {
                        "id": 7701,
                        "name": "Catalytic converter",
                        "description": "3-way catalytic converter",
                        "value": "standard",
                        "items": [
                            {
                                "id": 7702,
                                "name": "type",
                                "value": "3-way"
                            }
                        ]
                    },
                    {
                        "id": 8501,
                        "name": "Fuel system",
                        "description": "Multi-point injection fuel system",
                        "value": "standard",
                        "items": [
                            {
                                "id": 8502,
                                "name": "injection/carburation",
                                "value": "multi-point injection"
                            }
                        ]
                    },
                    {
                        "id": 8701,
                        "name": "Fuel",
                        "description": "Unleaded fuel",
                        "value": "standard",
                        "items": [
                            {
                                "id": 8702,
                                "name": "Fuel type",
                                "value": "unleaded"
                            },
                            {
                                "id": 8708,
                                "name": "generic primary fuel type",
                                "value": "petrol"
                            }
                        ]
                    },
                    {
                        "id": 13501,
                        "name": "Performance",
                        "description": "Performance: maximum speed (mph): 115, maximum speed (km/h): 185 and acceleration 0-100 km/h (secs): 10.9",
                        "value": "standard",
                        "items": [
                            {
                                "id": 13502,
                                "name": "maximum speed (km/h)",
                                "value": 185
                            },
                            {
                                "id": 13503,
                                "name": "acceleration 0-62mph (s)",
                                "value": 10.9
                            },
                            {
                                "id": 113502,
                                "name": "maximum speed (mph)",
                                "value": 115
                            }
                        ]
                    },
                    {
                        "id": 15301,
                        "name": "Power",
                        "description": "Power: 74 kW , 100 HP ISO @ 6,000 rpm; , 145 Nm @ 4,000 rpm",
                        "value": "standard",
                        "items": [
                            {
                                "id": 15302,
                                "name": "measurement standard",
                                "value": "ISO"
                            },
                            {
                                "id": 15303,
                                "name": "Maximum power kW",
                                "value": 74
                            },
                            {
                                "id": 15304,
                                "name": "Maximum power hp/PS",
                                "value": 100
                            },
                            {
                                "id": 15305,
                                "name": "rpm for maximum power (low)",
                                "value": 6000
                            },
                            {
                                "id": 15307,
                                "name": "maximum torque Nm",
                                "value": 145
                            },
                            {
                                "id": 15308,
                                "name": "rpm for maximum torque (low)",
                                "value": 4000
                            }
                        ]
                    },
                    {
                        "id": 15401,
                        "name": "Fuel consumption",
                        "description": "Fuel consumption EU 96 urban: 9.4, EU 96 std country: 5.7 and EU 96 std combined: 7",
                        "value": "standard",
                        "items": [
                            {
                                "id": 15409,
                                "name": "EU 96 std - urban (l/100km)",
                                "value": 9.4
                            },
                            {
                                "id": 15410,
                                "name": "EU 96 std - country (l/100km)",
                                "value": 5.7
                            },
                            {
                                "id": 15411,
                                "name": "EU 96 std - combined (l/100km)",
                                "value": 7
                            }
                        ]
                    },
                    {
                        "id": 18501,
                        "name": "Power steering",
                        "description": "Power steering",
                        "value": "standard"
                    },
                    {
                        "id": 20001,
                        "name": "Suspension",
                        "description": "Independent strut front suspension with anti-roll bar and coil springs , independent multi-link rear suspension with anti-roll bar and coil springs",
                        "value": "standard",
                        "items": [
                            {
                                "id": 20002,
                                "name": "type",
                                "location": "Front",
                                "value": "strut"
                            },
                            {
                                "id": 20002,
                                "name": "type",
                                "location": "Rear",
                                "value": "multi-link"
                            },
                            {
                                "id": 20003,
                                "name": "anti-roll bar",
                                "location": "Front",
                                "value": "yes"
                            },
                            {
                                "id": 20003,
                                "name": "anti-roll bar",
                                "location": "Rear",
                                "value": "yes"
                            },
                            {
                                "id": 20005,
                                "name": "wheel dependence",
                                "location": "Front",
                                "value": "independent"
                            },
                            {
                                "id": 20005,
                                "name": "wheel dependence",
                                "location": "Rear",
                                "value": "independent"
                            },
                            {
                                "id": 20006,
                                "name": "spring type",
                                "location": "Front",
                                "value": "coil"
                            },
                            {
                                "id": 20006,
                                "name": "spring type",
                                "location": "Rear",
                                "value": "coil"
                            }
                        ]
                    },
                    {
                        "id": 20601,
                        "name": "Transmission",
                        "description": "Manual five-speed transmission with gear lever on floor",
                        "value": "standard",
                        "items": [
                            {
                                "id": 20602,
                                "name": "Transmission type",
                                "value": "manual"
                            },
                            {
                                "id": 20603,
                                "name": "number of speeds",
                                "value": 5
                            },
                            {
                                "id": 20610,
                                "name": "gearchange location",
                                "value": "floor"
                            }
                        ]
                    },
                    {
                        "id": 38801,
                        "name": "Main service",
                        "description": "Main service distance 20,000 and period (mths) 12",
                        "value": "standard",
                        "items": [
                            {
                                "id": 38802,
                                "name": "distance (km)",
                                "value": 20000
                            },
                            {
                                "id": 38803,
                                "name": "period (mths)",
                                "value": 12
                            },
                            {
                                "id": 38805,
                                "name": "distance (miles)",
                                "value": 12427
                            }
                        ]
                    },
                    {
                        "id": 42001,
                        "name": "Fuel consumption",
                        "description": "Fuel consumption: EU 96 urban (l/100km): 9.4, country/highway (l/100km): 5.7 and combined (l/100km): 7",
                        "value": "standard",
                        "items": [
                            {
                                "id": 42003,
                                "name": "urban (l/100km)",
                                "value": 9.4
                            },
                            {
                                "id": 42004,
                                "name": "country/highway (l/100km)",
                                "value": 5.7
                            },
                            {
                                "id": 42005,
                                "name": "combined (l/100km)",
                                "value": 7
                            },
                            {
                                "id": 142003,
                                "name": "urban (mpg)",
                                "value": 25
                            },
                            {
                                "id": 142004,
                                "name": "country/highway (mpg)",
                                "value": 41
                            },
                            {
                                "id": 142005,
                                "name": "combined (mpg)",
                                "value": 34
                            }
                        ]
                    }
                ]
            },
            {
                "Others": [
                    {
                        "id": 174,
                        "name": "Global segment",
                        "description": "Global segment",
                        "value": "Lower Medium"
                    },
                    {
                        "id": 176,
                        "name": "Regional segment",
                        "description": "Regional segment",
                        "value": "C1 - lower medium -"
                    },
                    {
                        "id": 401,
                        "name": "Trim",
                        "description": "Trim level: ZETEC",
                        "value": "standard",
                        "items": [
                            {
                                "id": 402,
                                "name": "Trim level",
                                "value": "ZETEC"
                            },
                            {
                                "id": 404,
                                "name": "local trim level",
                                "value": "Zetec"
                            },
                            {
                                "id": 405,
                                "name": "trim classification",
                                "value": "S1"
                            }
                        ]
                    },
                    {
                        "id": 601,
                        "name": "Body style",
                        "description": "Five-door hatchback body style; RHD",
                        "value": "standard",
                        "items": [
                            {
                                "id": 602,
                                "name": "Number of doors",
                                "value": 5
                            },
                            {
                                "id": 603,
                                "name": "Body type",
                                "value": "hatchback"
                            },
                            {
                                "id": 605,
                                "name": "local number of doors",
                                "value": 5
                            },
                            {
                                "id": 606,
                                "name": "local body type",
                                "value": "hatchback"
                            },
                            {
                                "id": 609,
                                "name": "Driver location",
                                "value": "RHD"
                            }
                        ]
                    },
                    {
                        "id": 2501,
                        "name": "Insurance",
                        "description": "Insurance: 5E",
                        "value": "standard",
                        "items": [
                            {
                                "id": 2502,
                                "name": "description",
                                "value": "5E"
                            }
                        ]
                    },
                    {
                        "id": 3501,
                        "name": "Charges",
                        "description": "Charges: On Road Price, 12,245, 0 and 0",
                        "value": "standard",
                        "items": [
                            {
                                "id": 3510,
                                "name": "national tax 1 name",
                                "value": "On Road Price"
                            },
                            {
                                "id": 3516,
                                "name": "national tax 1 amount",
                                "value": 12245
                            },
                            {
                                "id": 3517,
                                "name": "national tax 2 amount",
                                "value": 0
                            },
                            {
                                "id": 3518,
                                "name": "national tax 3 amount",
                                "value": 0
                            }
                        ]
                    },
                    {
                        "id": 3601,
                        "name": "Delivery charges",
                        "description": "Standard delivery charges: 0",
                        "value": "standard",
                        "items": [
                            {
                                "id": 3602,
                                "name": "type",
                                "value": "standard"
                            },
                            {
                                "id": 3603,
                                "name": "amount",
                                "value": 0
                            }
                        ]
                    },
                    {
                        "id": 14601,
                        "name": "Cargo area cover/rear parcel shelf",
                        "description": "Rigid cargo area cover/rear parcel shelf",
                        "value": "standard",
                        "items": [
                            {
                                "id": 14602,
                                "name": "type",
                                "value": "rigid"
                            }
                        ]
                    },
                    {
                        "id": 23501,
                        "name": "Warranty whole vehicle - Total",
                        "description": "Full car warranty: duration (months): 12 or distance (miles): unlimited, distance (km): unlimited",
                        "value": "standard",
                        "items": [
                            {
                                "id": 23502,
                                "name": "duration (months)",
                                "value": 12
                            },
                            {
                                "id": 23503,
                                "name": "distance (km)",
                                "value": 999999
                            },
                            {
                                "id": 123503,
                                "name": "distance (miles)",
                                "value": 999999
                            }
                        ]
                    },
                    {
                        "id": 23601,
                        "name": "Warranty powertrain - Total",
                        "description": "Powertrain warranty: duration (months): 12 or distance (miles): unlimited, distance (km): unlimited",
                        "value": "standard",
                        "items": [
                            {
                                "id": 23602,
                                "name": "duration (months)",
                                "value": 12
                            },
                            {
                                "id": 23603,
                                "name": "distance (km)",
                                "value": 999999
                            },
                            {
                                "id": 123603,
                                "name": "distance (miles)",
                                "value": 999999
                            }
                        ]
                    },
                    {
                        "id": 23701,
                        "name": "Warranty anti-corrosion",
                        "description": "Anticorrosion warranty: duration (months): 144 or distance (miles): unlimited, distance (km): unlimited",
                        "value": "standard",
                        "items": [
                            {
                                "id": 23702,
                                "name": "duration (months)",
                                "value": 144
                            },
                            {
                                "id": 23703,
                                "name": "distance (km)",
                                "value": 999999
                            },
                            {
                                "id": 123703,
                                "name": "distance (miles)",
                                "value": 999999
                            }
                        ]
                    },
                    {
                        "id": 23801,
                        "name": "Warranty paint",
                        "description": "Paint warranty: duration (months): 12 or distance (miles): unlimited, distance (km): unlimited",
                        "value": "standard",
                        "items": [
                            {
                                "id": 23802,
                                "name": "duration (months)",
                                "value": 12
                            },
                            {
                                "id": 23803,
                                "name": "distance (km)",
                                "value": 999999
                            },
                            {
                                "id": 123803,
                                "name": "distance (miles)",
                                "value": 999999
                            }
                        ]
                    },
                    {
                        "id": 24001,
                        "name": "Warranty roadside assistance",
                        "description": "Road-side assistance warranty: duration (months): 12 or distance (miles): unlimited, distance (km): unlimited",
                        "value": "standard",
                        "items": [
                            {
                                "id": 24002,
                                "name": "duration (months)",
                                "value": 12
                            },
                            {
                                "id": 24003,
                                "name": "distance (km)",
                                "value": 999999
                            },
                            {
                                "id": 124003,
                                "name": "distance (miles)",
                                "value": 999999
                            }
                        ]
                    },
                    {
                        "id": 34201,
                        "name": "Check control",
                        "description": "Dash mounted vehicle warning system",
                        "value": "standard"
                    },
                    {
                        "id": 42901,
                        "name": "Date introduced",
                        "description": "Date introduced. Body type introduced: 19981015, Number of doors introduced: 19981015, Version introduced: 19981015 and Model introduced: 19981015",
                        "value": "standard",
                        "items": [
                            {
                                "id": 42906,
                                "name": "Body type introduced",
                                "value": 19981015
                            },
                            {
                                "id": 42907,
                                "name": "Number of doors introduced",
                                "value": 19981015
                            },
                            {
                                "id": 42908,
                                "name": "Version introduced",
                                "value": 19981015
                            },
                            {
                                "id": 42909,
                                "name": "Model introduced",
                                "value": 19981015
                            }
                        ]
                    }
                ]
            }
        ],
        "serviceVersion": "V2",
        "_links": {
            "self": {
                "href": "https://staging.motorspecs.com/specs/standard/1"
            }
        }
    }

    # print(dict['standardSpecification'][5])
    temp_dict = dict['standardSpecification']
    # print(temp_dict.get('Technical'))
    exact_item = 0
    for section in temp_dict:
        # print(section)
        if list(section.keys())[0] == 'Technical':
            # print('FOUND IT!!')
            break
        exact_item = exact_item + 1
    # print('Found at: ' + str(exact_item))

    tech_data = temp_dict[exact_item]
    print(tech_data)

    exact_item_2 = 0
    for item in tech_data['Technical']:
        current_item = (list(item.items()))
        # prints through technical items
        # print(current_item[0])
        if str(current_item[0]) == "('id', 7601)":
            emissions_items = current_item[4][1]
            # break
        if str(current_item[0]) == "('id', 42001)":
            combined_items = current_item[4][1]
            break
        exact_item_2 = exact_item_2 + 1

    if emissions_items[0]['id'] == 7602:
        try:
            eu_rating = emissions_items[0]['value']
            print(eu_rating)
        except IndexError:
            print('No EU status available from DVLA or Specs')
        try:
            co2_value = emissions_items[1]['value']
            print(co2_value)
        except IndexError:
            print('No CO2 value available from DVLA or Specs')
    elif emissions_items[0]['id'] == 7603:
        try:
            co2_value = emissions_items[0]['value']
            print(co2_value)
        except IndexError:
            print('No CO2 value available')
    else:
        print('No EU status or CO2 value available from DVLA or Specs')


def get_valuation(vrm, mileage, row):
    api_url = 'https://staging.motorspecs.com/valuation-glass/value'

    headers = {'Accept': 'application/vnd.valuation-glass.v2+json',
               'Content-Type': 'application/vnd.valuation-glass.v2+json',
               'Authorization': 'Bearer XXX'}

    body = {'registration': vrm,
            'currentMiles': mileage}

    response = requests.post(url=api_url, json=body, headers=headers)

    if response.status_code != 201:
        notice = ("No valuation available from Glass's")
        new_sheet.write(row, 20, notice)
        # print(notice)
    else:

        response_dict = json.loads(response.content)

        adjustedTradeValues = response_dict['glassValuation'][0]['adjustedTradeValues']

        # tradeValue = adjustedTradeValues['trade']
        retailValue = adjustedTradeValues['retail']
        # tradeHighValue = adjustedTradeValues['tradeHigh']
        # tradeAverageValue = adjustedTradeValues['tradeAverage']
        # tradeLowValue = adjustedTradeValues['tradeLow']

        # new_sheet.write(row, 4, tradeValue)
        new_sheet.write(row, 5, retailValue)
        # new_sheet.write(row, 6, tradeHighValue)
        # new_sheet.write(row, 7, tradeAverageValue)
        # new_sheet.write(row, 8, tradeLowValue)

        return

        # print(response_dict)


def test_val():
    data = '''
    {"vehicleId":null,"registration":"yg67uvd","currentMiles":50000,"priceWhenNew":28805,"glassValuation":[{"version":"116 1.5TD ( bhp ) ( s\/s ) Sports Hatch 2017.5MY","modelId":252296,"modelQualifier":"001","qualifiedModelCode":252296001,"glassCode":"48BH","newPrice":25810,"averageMileage":24000,"basicValue":{"trade":11950,"retail":14550},"adjustedMileage":50000,"adjustedTradeValues":{"trade":10900,"retail":13760,"tradeHigh":11080,"tradeAverage":10400,"tradeLow":10130},"adjustedConsumerValues":{"partExExcellent":10720,"partExAverage":9670,"partExLow":8560,"retail":13760,"privateSale":12440,"retailTransacted":13340},"commercialVehicle":{"lowMileageTrade":0,"lowMileageRetail":0,"disposalTrade":0}}],"serviceVersion":"V2","_links":{"self":{"href":"https:\/\/staging.motorspecs.com\/valuation-glass\/value"}}}
    '''
    response_dict = json.loads(data)

    adjustedTradeValues = response_dict['glassValuation'][0]['adjustedTradeValues']

    # tradeValue = adjustedTradeValues['trade']
    print(adjustedTradeValues)
    print(adjustedTradeValues['trade'])


if __name__ == "__main__":
    root = tk.Tk()
    root.title('MotorCheck')
    root.withdraw()
    file_path = tk.filedialog.askopenfilename()

    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)

    new_book = copy(book)
    new_sheet = new_book.get_sheet(0)
    new_sheet.write(0, 23, 'Notes')

    check_vrm()
    # get_euroStatus()
    # dict_Test()
