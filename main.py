import pandas as pd
import openpyxl
import math
import json

ignore_list = ['STATUS', 'INTERLOCKS', 'PDPs', 'PDP WORKSPACE', 'IP ADDRESS', 'Sheet1', '201', '203', '206', '401',
               '402', '403', '406', '611', '612', '621', '622', '631', '632', '633', '634', '901', '902', '903', '906', '931', '932', '933', '951', '952',
               '953', '956', '971', '981', '982', '983']

prefixes = ['SB+', 'RG+', 'BT+', 'OR+', 'RT+', 'SP+', 'RX+', 'SL+', 'IS+', 'BM+', 'RA+', 'BB+', 'SS+', 'AT+', 'CW+',
            'TP+', 'CE+', 'VB+', 'WD+', 'GU+', 'BS+', 'WG+', 'CH+', 'SC+', 'TL+', 'TU+', 'PL+', 'GC+']

eag_list = []
eag_list_size = 0
eag_ip_list = []

ecg_list = []
ecg_list_size = 0
ecg_ip_list = []

fmd_list = []
fmd_list_size = 0
fmd_ip_list = []

bkio_list = []
bkio_list_size = 0
bkio_ip_list = []

sol_list = []
sol_list_size = 0
sol_ip_list = []

unit_list_initial = []
unit_list_final = []


def do_stuff(sheet):
    print(f"Received: {sheet}")

    subnet = ''

    if sheet == '601':
        subnet = '1.'
    # elif sheet == '632':
    #     subnet = '2.'
    # elif sheet == '633':
    #     subnet = '3.'
    # elif sheet == '634':
    #     subnet = '4.'

    df = xl.parse(sheet)

    df1 = pd.DataFrame(df)

    names = df1.head()

    df_fmds = pd.DataFrame()

    for column in df1:
        data = df1[column]
        # print(data.name)

        if data.name == 'ENET1' or data.name == 'ENET2' or data.name == 'ENET3' or data.name == 'ENET4' or data.name == 'ENET5' or data.name == 'ENET6':

            for row in data:
                if 'EAG' in str(row):
                    eag_list.append([row, data.name])

                if 'ECG' in str(row):
                    ecg_list.append([row, data.name])

                if 'FMD' in str(row):
                    fmd_list.append([row, data.name])

                if 'BKIO' in str(row):
                    bkio_list.append([row, data.name])

                if 'SOL' in str(row):
                    sol_list.append([row, data.name])

    # START EAG IP
    eag_list.sort(key=lambda x: x[0])

    df_eags = pd.DataFrame(eag_list, columns=["NAME", "NETWORK"])

    eag_list_size = len(eag_list)

    ip_inc = 51

    for i in range(0, eag_list_size):
        eag_ip_list.append(subnet + str(ip_inc))
        ip_inc += 1

    df_eags["IP"] = eag_ip_list

    # START ECG IP
    ecg_list.sort(key=lambda x: x[0])

    df_ecgs = pd.DataFrame(ecg_list, columns=["NAME", "NETWORK"])

    ecg_list_size = len(ecg_list)

    ip_inc = 61

    for i in range(0, ecg_list_size):
        ecg_ip_list.append(subnet + str(ip_inc))
        ip_inc += 1

    df_ecgs["IP"] = ecg_ip_list

    # START FMD IP
    fmd_list.sort(key=lambda x: x[0])

    df_fmds = pd.DataFrame(fmd_list, columns=["NAME", "NETWORK"])

    fmd_list_size = len(fmd_list)

    ip_inc = 101

    for i in range(0, fmd_list_size):
        fmd_ip_list.append(subnet + str(ip_inc))
        ip_inc += 1

    df_fmds["IP"] = fmd_ip_list

    # START BKIO IP
    bkio_list.sort(key=lambda x: x[0])

    df_bkio = pd.DataFrame(bkio_list, columns=["NAME", "NETWORK"])

    bkio_list_size = len(bkio_list)

    ip_inc = 201

    for i in range(0, bkio_list_size):
        bkio_ip_list.append(subnet + str(ip_inc))
        ip_inc += 1

    df_bkio["IP"] = bkio_ip_list

    # START SOL IP
    sol_list.sort(key=lambda x: x[0])

    df_sol = pd.DataFrame(sol_list, columns=["NAME", "NETWORK"])

    sol_list_size = len(sol_list)

    ip_inc = 231

    for i in range(0, sol_list_size):
        sol_ip_list.append(subnet + str(ip_inc))
        ip_inc += 1

    df_sol["IP"] = sol_ip_list

    # START MERGE AND OUTPUT
    df_merge = pd.concat([df_eags, df_ecgs, df_fmds, df_bkio, df_sol])
    # print(df_merge)
    # df_merge.to_excel('output.xlsx', sheet_name=sheet)

    # SORT AND FORMAT
    df_merge.sort_values(by=['NETWORK', 'IP'], ascending=True, inplace=True)
    print(df_merge)
    df_merge.to_excel(f"output_{sheet}.xlsx", sheet_name=sheet)


if __name__ == "__main__":
    print('main')

    xl = pd.ExcelFile('ENET_LAYOUT.xlsx')
    for sheet in xl.sheet_names:
        if sheet in ignore_list:
            continue

        # SET DEFAULTS FOR EACH ITERATION
        eag_list = []
        eag_list_size = 0
        eag_ip_list = []

        ecg_list = []
        ecg_list_size = 0
        ecg_ip_list = []

        fmd_list = []
        fmd_list_size = 0
        fmd_ip_list = []

        bkio_list = []
        bkio_list_size = 0
        bkio_ip_list = []

        sol_list = []
        sol_list_size = 0
        sol_ip_list = []

        do_stuff(sheet)

    # df_input = pd.read_excel('output_621.xlsx', index_col=0)
    # df = pd.DataFrame(df_input, columns=["NAME", "NETWORK", "IP"])
    # # print(df)
    #
    # for column in df:
    #     data = df[column]
    #
    #     if data.name == 'NAME':
    #         for row in data:
    #             if row[0] == 'U':
    #                 unit = row
    #                 unit = unit[1:]
    #                 unit_list_initial.append(unit)
    #
    #
    # print(unit_list_initial)
    #
    # file = open('Test.json')
    # json_obj = json.load(file)
    #
    # unit_counts = 0
    #
    # for snapper in json_obj['snappers']:
    #     print(snapper)
    #     if 'unitID' in snapper:
    #         unit_counts += 1
    #         if any(pre in snapper['unitID'] for pre in prefixes):
    #             # print(snapper['unitID'])
    #             unitID = snapper['unitID']
    #             deviceID = snapper['deviceId']
    #             print(unitID)
    #         elif any(pre in snapper['deviceId'] for pre in prefixes):
    #             # print(snapper['unitID'])
    #             deviceID = snapper['deviceId']
    #             print(deviceID)
    #
    #
    # print(f"{unit_counts} total units in provided json file...")
