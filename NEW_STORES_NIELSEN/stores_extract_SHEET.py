from stores_conditions import *
from PyQt5 import QtWidgets as qtw
import sys

# Notepad directory
txt_path = 'C:/Users/tsst1002/Desktop/STORES_LOAD/'

# Notepad file names
stores_txt = '1.stores.txt'
dtgroups_txt = '2.dtgroups.txt'
xcodegroups_txt = '3.xcodegroups.txt'
lbatchstores_txt = '4.lbatchstores.txt'
store_group_stores_txt = '5.store_group_stores.txt'
lbatches_txt = '6.lbatches.txt'
lbatchdtgroups_txt = '7.lbatchdtgroups.txt'
spell_check_txt = '8.SPELLCHECK.txt'

# Input variables
country_input = ''
excel_name_input = ''
sheet_num_input = ''
start_input = ''
end_input = ''


def stores_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Shoptypes variable
    shop_type = shoptype(country.upper())

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Columns definition
            if country.upper() == 'KZ':
                ac_nshopid = {
                    'SCANNING': f'KZ00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'KZM0{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}')
                }
                ac_shopstatus = 'REQUIRED'
                ac_defaultxcodegr = {
                    'SCANNING': sheet.cell(row=row, column=6).value,
                    'AUDIT': sheet.cell(row=row, column=4).value,
                    'A2S': sheet.cell(row=row, column=6).value
                }
                ac_countryid = country.upper()
                ac_languageid = 'EN'
                nc_acv = 100
                nc_activeflag = 1
                nc_dupitems_flag = 0
                nc_eanxcode_flag = (0, 1)
                ac_channelid = {
                    'SCANNING': sheet.cell(row=row, column=7).value,
                    'AUDIT': sheet.cell(row=row, column=5).value,
                    'A2S': (sheet.cell(row=row, column=7).value, 'DUMMY')
                }
                nc_dummy_flag = (0, 1)
                ac_shopdescription = {
                    'SCANNING': sheet.cell(row=row, column=8).value,
                    'AUDIT': sheet.cell(row=row, column=6).value,
                    'A2S': sheet.cell(row=row, column=8).value
                }
                ac_retailer = {
                    'SCANNING': sheet.cell(row=row, column=9).value,
                    'AUDIT': 'AUDIT',
                    'A2S': sheet.cell(row=row, column=9).value
                }
                ac_area = {
                    'SCANNING': sheet.cell(row=row, column=10).value,
                    'AUDIT': sheet.cell(row=row, column=7).value,
                    'A2S': sheet.cell(row=row, column=10).value
                }
                ac_shoptype = {
                    'SCANNING': shop_type.get(sheet.cell(row=row, column=10).value),
                    'AUDIT': 'AUDIT',
                    'A2S': shop_type.get(sheet.cell(row=row, column=10).value)
                }
                nc_surface = {
                    'SCANNING': sheet.cell(row=row, column=11).value,
                    'AUDIT': sheet.cell(row=row, column=8).value,
                    'A2S': sheet.cell(row=row, column=11).value
                }
                nc_xf = 1
                ac_store_char1 = {
                    'SCANNING': sheet.cell(row=row, column=12).value,
                    'AUDIT': sheet.cell(row=row, column=9).value,
                    'A2S': sheet.cell(row=row, column=12).value
                }
            elif country.upper() == 'BY':
                ac_nshopid = {
                    'SCANNING': f'BY00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'BY00{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'BYM0{sheet.cell(row=row, column=5).value}', f'BYAS{sheet.cell(row=row, column=5).value}')
                }
                ac_shopstatus = 'REQUIRED'
                ac_defaultxcodegr = {
                    'SCANNING': sheet.cell(row=row, column=6).value,
                    'AUDIT': sheet.cell(row=row, column=4).value,
                    'A2S': sheet.cell(row=row, column=6).value
                }
                ac_countryid = country.upper()
                ac_languageid = 'EN'
                nc_acv = 100
                nc_activeflag = 1
                nc_dupitems_flag = 0
                nc_eanxcode_flag = (0, 1)
                ac_channelid = {
                    'SCANNING': sheet.cell(row=row, column=7).value,
                    'AUDIT': sheet.cell(row=row, column=5).value,
                    'A2S': (sheet.cell(row=row, column=7).value, 'DUMMY')
                }
                nc_dummy_flag = (0, 1)
                ac_shopdescription = {
                    'SCANNING': sheet.cell(row=row, column=8).value,
                    'AUDIT': sheet.cell(row=row, column=6).value,
                    'A2S': sheet.cell(row=row, column=8).value
                }
                ac_retailer = {
                    'SCANNING': sheet.cell(row=row, column=9).value,
                    'AUDIT': 'AUDIT',
                    'A2S': sheet.cell(row=row, column=9).value
                }
                ac_area = {
                    'SCANNING': sheet.cell(row=row, column=10).value,
                    'AUDIT': sheet.cell(row=row, column=7).value,
                    'A2S': sheet.cell(row=row, column=10).value
                }
                ac_shoptype = {
                    'SCANNING': shop_type.get(sheet.cell(row=row, column=10).value),
                    'AUDIT': 'AUDIT',
                    'A2S': shop_type.get(sheet.cell(row=row, column=10).value)
                }
                nc_surface = {
                    'SCANNING': sheet.cell(row=row, column=11).value,
                    'AUDIT': sheet.cell(row=row, column=8).value,
                    'A2S': sheet.cell(row=row, column=11).value
                }
                nc_xf = 1
                ac_store_char1 = {
                    'SCANNING': sheet.cell(row=row, column=12).value,
                    'AUDIT': sheet.cell(row=row, column=9).value,
                    'A2S': sheet.cell(row=row, column=12).value
                }

            # Columns extraction order
            if sheet_name.upper() == 'SCANNING':
                file.write(
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_shopstatus}\t'
                    f'{ac_defaultxcodegr["SCANNING"]}\t'
                    f'{ac_countryid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_acv}\t'
                    f'{nc_activeflag}\t'
                    f'{nc_dupitems_flag}\t'
                    f'{nc_eanxcode_flag[0]}\t'
                    f'{ac_channelid["SCANNING"]}\t'
                    f'{nc_dummy_flag[0]}\t'
                    f'{ac_shopdescription["SCANNING"]}\t'
                    f'{ac_retailer["SCANNING"]}\t'
                    f'{ac_area["SCANNING"]}\t'
                    f'{ac_shoptype["SCANNING"]}\t'
                    f'{nc_surface["SCANNING"]}\t'
                    f'{nc_xf}\t'
                    f'{ac_store_char1["SCANNING"]}\n'
                )
            elif sheet_name.upper() == 'AUDIT':
                file.write(
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_shopstatus}\t'
                    f'{ac_defaultxcodegr["AUDIT"]}\t'
                    f'{ac_countryid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_acv}\t'
                    f'{nc_activeflag}\t'
                    f'{nc_dupitems_flag}\t'
                    f'{nc_eanxcode_flag[1]}\t'
                    f'{ac_channelid["AUDIT"]}\t'
                    f'{nc_dummy_flag[0]}\t'
                    f'{ac_shopdescription["AUDIT"]}\t'
                    f'{ac_retailer["AUDIT"]}\t'
                    f'{ac_area["AUDIT"]}\t'
                    f'{ac_shoptype["AUDIT"]}\t'
                    f'{nc_surface["AUDIT"]}\t'
                    f'{nc_xf}\t'
                    f'{ac_store_char1["AUDIT"]}\n'
                )
            elif sheet_name.upper() == 'A2S':
                file.write(
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_shopstatus}\t'
                    f'{ac_defaultxcodegr["A2S"]}\t'
                    f'{ac_countryid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_acv}\t'
                    f'{nc_activeflag}\t'
                    f'{nc_dupitems_flag}\t'
                    f'{nc_eanxcode_flag[0]}\t'
                    f'{ac_channelid["A2S"][0]}\t'
                    f'{nc_dummy_flag[0]}\t'
                    f'{ac_shopdescription["A2S"]}\t'
                    f'{ac_retailer["A2S"]}\t'
                    f'{ac_area["A2S"]}\t'
                    f'{ac_shoptype["A2S"]}\t'
                    f'{nc_surface["A2S"]}\t'
                    f'{nc_xf}\t'
                    f'{ac_store_char1["A2S"]}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_shopstatus}\t'
                    f'{ac_defaultxcodegr["A2S"]}\t'
                    f'{ac_countryid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_acv}\t'
                    f'{nc_activeflag}\t'
                    f'{nc_dupitems_flag}\t'
                    f'{nc_eanxcode_flag[1]}\t'
                    f'{ac_channelid["A2S"][1]}\t'
                    f'{nc_dummy_flag[1]}\t'
                    f'{ac_shopdescription["A2S"]}\t'
                    f'{ac_retailer["A2S"]}\t'
                    f'{ac_area["A2S"]}\t'
                    f'{ac_shoptype["A2S"]}\t'
                    f'{nc_surface["A2S"]}\t'
                    f'{nc_xf}\t'
                    f'{ac_store_char1["A2S"]}\n'
                )

    none = none_check(txt_path, txt_name)
    print(f'Data extracted to {txt_name}{none}')


def storedtgroups_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Columns definition
            if country.upper() == 'KZ':
                ac_nshopid = {
                    'SCANNING': f'KZ00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'KZM0{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}')
                }
                ac_dtgroup = {
                    'SCANNING': 'VOLUMETRIC',
                    'AUDIT': ('AUDIT_ORIG', 'AUDIT_DTYPE', 'CAUSA1'),
                    'A2S': ('AUDIT_DTYPE', 'VOLUMETRIC')
                }
                nc_dtgroup = 1
                nc_activeflag = 1
            elif country.upper() == 'BY':
                ac_nshopid = {
                    'SCANNING': f'BY00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'BY00{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'BY00{sheet.cell(row=row, column=5).value}', f'BYAS{sheet.cell(row=row, column=5).value}')
                }
                ac_dtgroup = {
                    'SCANNING': 'VOLUMETRIC',
                    'AUDIT': ('AUDIT_ORIG', 'AUDIT_DTYPE'),
                    'A2S': ('AUDIT_DTYPE', 'VOLUMETRIC')
                }
                nc_dtgroup = 1
                nc_activeflag = 1

            # Columns extraction order
            if sheet_name.upper() == 'SCANNING':
                file.write(
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_dtgroup["SCANNING"]}\t'
                    f'{nc_dtgroup}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'AUDIT':
                file.write(
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_dtgroup["AUDIT"][0]}\t'
                    f'{nc_dtgroup}\t'
                    f'{nc_activeflag}\n'
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_dtgroup["AUDIT"][1]}\t'
                    f'{nc_dtgroup}\t'
                    f'{nc_activeflag}\n'
                )
                if country.upper() == 'KZ':
                    file.write(
                        f'{ac_nshopid["AUDIT"]}\t'
                        f'{ac_dtgroup["AUDIT"][2]}\t'
                        f'{nc_dtgroup}\t'
                        f'{nc_activeflag}\n'
                    )
            elif sheet_name.upper() == 'A2S':
                file.write(
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_dtgroup["A2S"][0]}\t'
                    f'{nc_dtgroup}\t'
                    f'{nc_activeflag}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_dtgroup["A2S"][1]}\t'
                    f'{nc_dtgroup}\t'
                    f'{nc_activeflag}\n'
                )

    none = none_check(txt_path, txt_name)
    print(f'Data extracted to {txt_name}{none}')


def storexcodegroups_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Columns definition
            if country.upper() == 'KZ':
                ac_nshopid = {
                    'SCANNING': f'KZ00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'KZM0{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}')
                }
                ac_xcodegr = {
                    'SCANNING': {
                        1: f'KZ00{sheet.cell(row=row, column=5).value}',
                        2: 'KZXX',
                        3: sheet.cell(row=row, column=6).value,
                        4: 'EAN'
                    },
                    'AUDIT': {
                        1: f'KZM0{sheet.cell(row=row, column=3).value}',
                        2: 'NAN_KEY',
                        3: 'KZ100_M'
                    },
                    'A2S': {
                        1: (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}'),
                        2: 'KZXX',
                        3: (sheet.cell(row=row, column=6).value, f'{sheet.cell(row=row, column=6).value}_M'),
                        4: ('EAN', 'NAN_KEY')
                    }
                }
                nc_xcodegrseq = {1: 1, 2: 2, 3: 3, 4: 4}
                nc_peractivefrom = 0
                nc_peractiveto = 9999
            elif country.upper() == 'BY':
                ac_nshopid = {
                    'SCANNING': f'BY00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'BY00{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'BY00{sheet.cell(row=row, column=5).value}', f'BYAS{sheet.cell(row=row, column=5).value}')
                }
                ac_xcodegr = {
                    'SCANNING': {
                        1: f'BY00{sheet.cell(row=row, column=5).value}',
                        2: 'BYXX',
                        3: sheet.cell(row=row, column=6).value,
                        4: 'EAN'
                    },
                    'AUDIT': {
                        1: f'BY00{sheet.cell(row=row, column=3).value}',
                        2: 'NAN_KEY',
                        3: 'BY100_M'
                    },
                    'A2S': {
                        1: (f'BY00{sheet.cell(row=row, column=5).value}', f'BYAS{sheet.cell(row=row, column=5).value}'),
                        2: 'BYXX',
                        3: (sheet.cell(row=row, column=6).value, f'{sheet.cell(row=row, column=6).value}_M'),
                        4: ('EAN', 'NAN_KEY')
                    }
                }
                nc_xcodegrseq = {1: 1, 2: 2, 3: 3, 4: 4}
                nc_peractivefrom = 0
                nc_peractiveto = 9999

            # Columns extraction order
            if sheet_name.upper() == 'SCANNING':
                file.write(
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_xcodegr["SCANNING"][1]}\t'
                    f'{nc_xcodegrseq[1]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_xcodegr["SCANNING"][2]}\t'
                    f'{nc_xcodegrseq[2]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_xcodegr["SCANNING"][3]}\t'
                    f'{nc_xcodegrseq[3]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["SCANNING"]}\t'
                    f'{ac_xcodegr["SCANNING"][4]}\t'
                    f'{nc_xcodegrseq[4]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                )
            elif sheet_name.upper() == 'AUDIT':
                file.write(
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_xcodegr["AUDIT"][1]}\t'
                    f'{nc_xcodegrseq[1]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_xcodegr["AUDIT"][2]}\t'
                    f'{nc_xcodegrseq[2]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{ac_xcodegr["AUDIT"][3]}\t'
                    f'{nc_xcodegrseq[3]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                )
            elif sheet_name.upper() == 'A2S':
                file.write(
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_xcodegr["A2S"][1][0]}\t'
                    f'{nc_xcodegrseq[1]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_xcodegr["A2S"][2]}\t'
                    f'{nc_xcodegrseq[2]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_xcodegr["A2S"][3][0]}\t'
                    f'{nc_xcodegrseq[3]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{ac_xcodegr["A2S"][4][0]}\t'
                    f'{nc_xcodegrseq[4]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_xcodegr["A2S"][1][1]}\t'
                    f'{nc_xcodegrseq[1]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_xcodegr["A2S"][2]}\t'
                    f'{nc_xcodegrseq[2]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_xcodegr["A2S"][3][1]}\t'
                    f'{nc_xcodegrseq[3]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{ac_xcodegr["A2S"][4][1]}\t'
                    f'{nc_xcodegrseq[4]}\t'
                    f'{nc_peractivefrom}\t'
                    f'{nc_peractiveto}\n'
                )

    none = none_check(txt_path, txt_name)
    print(f'Data extracted to {txt_name}{none}')


# AUDIT LHHT!
def lbatchstores_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Columns definition
            if country.upper() == 'KZ':
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': 'ACNIELSEN_MNTL',
                    'A2S': (sheet.cell(row=row, column=2).value, 'TTR_AGGREGATION')
                }
                ac_nshopid = {
                    'SCANNING': f'KZ00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'KZM0{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}')
                }
                ac_cshopid = {
                    'SCANNING': sheet.cell(row=row, column=5).value,
                    'AUDIT': sheet.cell(row=row, column=3).value,
                    'A2S': (sheet.cell(row=row, column=5).value, ac_nshopid['A2S'][0])
                }
                nc_cshopidreq = 0
                nc_activeflag = 1
            elif country.upper() == 'BY':
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': 'ACNIELSEN_MNTL',
                    'A2S': (sheet.cell(row=row, column=2).value,
                            'SISTERSTORES' if sheet.cell(row=row, column=7).value == 'SCAN' else 'SISTERSTORES_MNTL')
                }
                ac_nshopid = {
                    'SCANNING': f'BY00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'BY00{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'BY00{sheet.cell(row=row, column=5).value}', f'BYAS{sheet.cell(row=row, column=5).value}')
                }
                ac_cshopid = {
                    'SCANNING': sheet.cell(row=row, column=5).value,
                    'AUDIT': sheet.cell(row=row, column=3).value,
                    'A2S': (sheet.cell(row=row, column=5).value, ac_nshopid['A2S'][0])
                }
                nc_cshopidreq = 0
                nc_activeflag = 1

            # Columns extraction order BAU
            # if sheet_name.upper() == 'SCANNING':
            #     file.write(
            #         f'{ac_lbatchid["SCANNING"]}\t'
            #         f'{ac_cshopid["SCANNING"]}\t'
            #         f'{ac_nshopid["SCANNING"]}\t'
            #         f'{nc_cshopidreq}\t'
            #         f'{nc_activeflag}\n'
            #     )
            # elif sheet_name.upper() == 'AUDIT':
            #     file.write(
            #         f'{ac_lbatchid["AUDIT"]}\t'
            #         f'{ac_cshopid["AUDIT"]}\t'
            #         f'{ac_nshopid["AUDIT"]}\t'
            #         f'{nc_cshopidreq}\t'
            #         f'{nc_activeflag}\n'
            #     )
            # elif sheet_name.upper() == 'A2S':
            #     file.write(
            #         f'{ac_lbatchid["A2S"][0]}\t'
            #         f'{ac_cshopid["A2S"][0]}\t'
            #         f'{ac_nshopid["A2S"][1]}\t'
            #         f'{nc_cshopidreq}\t'
            #         f'{nc_activeflag}\n'
            #         f'{ac_lbatchid["A2S"][1]}\t'
            #         f'{ac_cshopid["A2S"][1]}\t'
            #         f'{ac_nshopid["A2S"][1]}\t'
            #         f'{nc_cshopidreq}\t'
            #         f'{nc_activeflag}\n'
            #     )

            # TRANSITION ONLY
            if sheet_name.upper() == 'SCANNING':
                file.write(
                        f'{ac_lbatchid["SCANNING"]}\t'
                        f'{ac_cshopid["SCANNING"]}\t'
                        f'{ac_nshopid["SCANNING"]}\t'
                        f'{nc_cshopidreq}\t'
                        f'{nc_activeflag}\n'
                        f'VOL1\t'
                        f'{ac_cshopid["SCANNING"]}\t'
                        f'{ac_nshopid["SCANNING"]}\t'
                        f'{nc_cshopidreq}\t'
                        f'{nc_activeflag}\n'
                    )
            elif sheet_name.upper() == 'AUDIT':
                file.write(
                    f'ACNIELSEN_MNTL\t'
                    f'{ac_cshopid["AUDIT"]}\t'
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                    f'VPSMONT\t'
                    f'{ac_cshopid["AUDIT"]}\t'
                    f'{ac_nshopid["AUDIT"]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'A2S':
                file.write(
                    f'{ac_lbatchid["A2S"][0]}\t'
                    f'{ac_cshopid["A2S"][0]}\t'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                    f'{ac_lbatchid["A2S"][1]}\t'
                    f'{ac_cshopid["A2S"][1]}\t'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                    f'VOL1\t'
                    f'{ac_cshopid["A2S"][0]}\t'
                    f'{ac_nshopid["A2S"][1]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                    f'VPSMONT\t'
                    f'{ac_cshopid["A2S"][0]}\t'
                    f'{ac_nshopid["A2S"][0]}\t'
                    f'{nc_cshopidreq}\t'
                    f'{nc_activeflag}\n'
                )

    none = none_check(txt_path, txt_name)
    print(f'Data extracted to {txt_name}{none}')


def store_group_stores(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Columns definition
            if country.upper() == 'KZ':
                ac_store_group = {
                    'SCANNING': sheet.cell(row=row, column=13).value,
                    'AUDIT': 'ZZZ',
                    'A2S': sheet.cell(row=row, column=13).value
                }
                ac_nshopid = {
                    'SCANNING': f'KZ00{sheet.cell(row=row, column=5).value}',
                    'AUDIT': f'KZM0{sheet.cell(row=row, column=3).value}',
                    'A2S': (f'KZM0{sheet.cell(row=row, column=5).value}', f'KZAS{sheet.cell(row=row, column=5).value}')
                }

                # Columns extraction order
                if sheet_name.upper() == 'SCANNING':
                    file.write(
                        f'{ac_store_group["SCANNING"]}\t'
                        f'{ac_nshopid["SCANNING"]}\n'
                    )
                elif sheet_name.upper() == 'AUDIT':
                    file.write(
                        f'{ac_store_group["AUDIT"]}\t'
                        f'{ac_nshopid["AUDIT"]}\n'
                    )
                elif sheet_name.upper() == 'A2S':
                    file.write(
                        f'{ac_store_group["A2S"]}\t'
                        f'{ac_nshopid["A2S"][0]}\n'
                        f'{ac_store_group["A2S"]}\t'
                        f'{ac_nshopid["A2S"][1]}\n'
                    )

    none = none_check(txt_path, txt_name)
    print(f'Data extracted to {txt_name}{none}')


def lbatches_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    new_batch = False

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Check for new batches for KZ
            if country.upper() == 'KZ' and \
            (sheet_name.upper() == 'SCANNING' and 'batch' in str(sheet.cell(row=row, column=14).value).lower()) or \
            (sheet_name.upper() == 'AUDIT' and 'batch' in str(sheet.cell(row=row, column=10).value).lower()) or \
            (sheet_name.upper() == 'A2S' and 'batch' in str(sheet.cell(row=row, column=14).value).lower()):

                new_batch = True
                # Columns definition
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': f'PHHTKZMM{sheet.cell(row=row, column=3).value}',
                    'A2S': sheet.cell(row=row, column=2).value
                }
                nc_activeflag = 1
            # Check for new batches for BY
            elif country.upper() == 'BY' and \
            (sheet_name.upper() == 'SCANNING' and 'batch' in str(sheet.cell(row=row, column=13).value).lower()) or \
            (sheet_name.upper() == 'AUDIT' and 'batch' in str(sheet.cell(row=row, column=10).value).lower()) or \
            (sheet_name.upper() == 'A2S' and 'batch' in str(sheet.cell(row=row, column=13).value).lower()):

                new_batch = True
                # Columns definition
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': f'PHHTBYMM{sheet.cell(row=row, column=3).value}',
                    'A2S': sheet.cell(row=row, column=2).value
                }
                nc_activeflag = 1

            # Columns extraction order
            if sheet_name.upper() == 'SCANNING' and new_batch:
                file.write(
                    f'{ac_lbatchid["SCANNING"]}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'AUDIT' and new_batch:
                file.write(
                    f'{ac_lbatchid["AUDIT"]}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'A2S' and new_batch:
                file.write(
                    f'{ac_lbatchid["A2S"]}\t'
                    f'{nc_activeflag}\n'
                )

    if new_batch:
        none = none_check(txt_path, txt_name)
        print(f'Data extracted to {txt_name}{none}')


def lbatchdtgroups_extract(country, txt_name, excel_name, sheet_num, start, end):

    # Directory function - storing the paths in variables
    notepad = notepad_directory(txt_path, txt_name)
    excel = excel_directory(country, excel_name)
    sheet = open_xl(excel)[0][sheet_num]
    sheet_name = open_xl(excel)[1][sheet_num]

    new_batch = False

    with open(notepad, 'w') as file:

        # .xlsx rows cycle starting from row 3
        for row in range(start, end + start):

            # Check for new batches for KZ
            if country.upper() == 'KZ' and \
            (sheet_name.upper() == 'SCANNING' and 'batch' in str(sheet.cell(row=row, column=14).value).lower()) or \
            (sheet_name.upper() == 'AUDIT' and 'batch' in str(sheet.cell(row=row, column=10).value).lower()) or \
            (sheet_name.upper() == 'A2S' and 'batch' in str(sheet.cell(row=row, column=14).value).lower()):

                new_batch = True
                # Columns definition
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': f'PHHTKZMM{sheet.cell(row=row, column=3).value}',
                    'A2S': sheet.cell(row=row, column=2).value
                }
                ac_dtgroup = {
                    'SCANNING': 'VOLUMETRIC',
                    'AUDIT': ('AUDIT_DTYPE', 'AUDIT_ORIG'),
                    'A2S': 'VOLUMETRIC'
                }
                nc_dtgroupseq = 1
                ac_currencyid = 'KZT'
                ac_languageid = 'EN'
                nc_activeflag = 1
            # Check for new batches for BY
            elif country.upper() == 'BY' and \
            (sheet_name.upper() == 'SCANNING' and 'batch' in str(sheet.cell(row=row, column=13).value).lower()) or \
            (sheet_name.upper() == 'AUDIT' and 'batch' in str(sheet.cell(row=row, column=10).value).lower()) or \
            (sheet_name.upper() == 'A2S' and 'batch' in str(sheet.cell(row=row, column=13).value).lower()):

                new_batch = True
                # Columns definition
                ac_lbatchid = {
                    'SCANNING': sheet.cell(row=row, column=2).value,
                    'AUDIT': f'PHHTBYMM{sheet.cell(row=row, column=3).value}',
                    'A2S': sheet.cell(row=row, column=2).value
                }
                ac_dtgroup = {
                    'SCANNING': 'VOLUMETRIC',
                    'AUDIT': ('AUDIT_DTYPE', 'AUDIT_ORIG'),
                    'A2S': 'VOLUMETRIC'
                }
                nc_dtgroupseq = 1
                ac_currencyid = 'BYB'
                ac_languageid = 'EN'
                nc_activeflag = 1

            # Columns extraction order
            if sheet_name.upper() == 'SCANNING' and new_batch:
                file.write(
                    f'{ac_lbatchid["SCANNING"]}\t'
                    f'{ac_dtgroup["SCANNING"]}\t'
                    f'{nc_dtgroupseq}\t'
                    f'{ac_currencyid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'AUDIT' and new_batch:
                file.write(
                    f'{ac_lbatchid["AUDIT"]}\t'
                    f'{ac_dtgroup["AUDIT"][0]}\t'
                    f'{nc_dtgroupseq}\t'
                    f'{ac_currencyid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_activeflag}\n'
                    f'{ac_lbatchid["AUDIT"]}\t'
                    f'{ac_dtgroup["AUDIT"][1]}\t'
                    f'{nc_dtgroupseq}\t'
                    f'{ac_currencyid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_activeflag}\n'
                )
            elif sheet_name.upper() == 'A2S' and new_batch:
                file.write(
                    f'{ac_lbatchid["A2S"]}\t'
                    f'{ac_dtgroup["A2S"]}\t'
                    f'{nc_dtgroupseq}\t'
                    f'{ac_currencyid}\t'
                    f'{ac_languageid}\t'
                    f'{nc_activeflag}\n'
                )

    if new_batch:
        none = none_check(txt_path, txt_name)
        print(f'Data extracted to {txt_name}{none}')


class MainWindow(qtw.QWidget):

    # Class constructor
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Main window properties
        self.leCountry = qtw.QLineEdit(self)
        self.leExcel = qtw.QLineEdit(self)
        self.leSheet = qtw.QLineEdit(self)
        self.leStart = qtw.QLineEdit(self)
        self.leEnd = qtw.QLineEdit(self)
        self.form_groupbox = qtw.QGroupBox('New Stores')
        self.form_layout = qtw.QFormLayout(self)

        # Create UI
        self.setup_UI()

        # Return line edit values on ENTER press
        self.leCountry.returnPressed.connect(self.inputs)
        self.leExcel.returnPressed.connect(self.inputs)
        self.leSheet.returnPressed.connect(self.inputs)
        self.leStart.returnPressed.connect(self.inputs)
        self.leEnd.returnPressed.connect(self.inputs)
        # Close UI window on ENTER press
        self.leCountry.returnPressed.connect(self.close)
        self.leExcel.returnPressed.connect(self.close)
        self.leSheet.returnPressed.connect(self.close)
        self.leStart.returnPressed.connect(self.close)
        self.leEnd.returnPressed.connect(self.close)
        # OK button click
        self.btn_Ok.clicked.connect(self.inputs)
        self.btn_Ok.clicked.connect(self.close)
        # Cancel button click
        self.btn_Cancel.clicked.connect(self.on_cancel)
        self.btn_Cancel.clicked.connect(self.close)

        # Show UI function
        self.show()

    # Create UI
    def setup_UI(self):
        self.setWindowTitle('Signals And Slots')
        self.setGeometry(500, 300, 600, 500)

        self.create_form_groupbox()
        self.create_buttons()

        # Create main Layout
        main_layout = qtw.QVBoxLayout(self)
        main_layout.addWidget(self.form_groupbox)
        main_layout.addLayout(self.buttons_layout)

    # Create Form Group Box
    def create_form_groupbox(self):
        self.form_groupbox.setLayout(self.form_layout)

        self.form_layout.addRow('Country code:', self.leCountry)
        self.form_layout.addRow('Excel file name:', self.leExcel)
        self.form_layout.addRow('Sheet (starts from 0):', self.leSheet)
        self.form_layout.addRow('Start row:', self.leStart)
        self.form_layout.addRow('End row:', self.leEnd)

    # Create Buttons Layout
    def create_buttons(self):
        self.buttons_layout = qtw.QHBoxLayout()
        self.btn_Ok = qtw.QPushButton('OK')
        self.btn_Cancel = qtw.QPushButton('Cancel')
        self.buttons_layout.addWidget(self.btn_Ok)
        self.buttons_layout.addWidget(self.btn_Cancel)

    # Collect inputs in variables
    def inputs(self):
        global country_input
        global excel_name_input
        global sheet_num_input
        global start_input
        global end_input

        country_input = self.leCountry.text()
        excel_name_input = f'{self.leExcel.text()}.xlsx'
        sheet_num_input = int(self.leSheet.text())
        start_input = int(self.leStart.text())
        end = int(self.leEnd.text())

        end_input = end - start_input + 1

    # Close message
    def on_cancel(self):
        print('Canceled!')


# Run extraction functions
if __name__ == '__main__':
    try:
        app = qtw.QApplication(sys.argv)
        window = MainWindow()
        app.exec_()

        # Spell check
        if spell_check_sheet(country_input, txt_path, spell_check_txt, excel_name_input, start_input, sheet_num_input):
            # Stores extraction
            stores_extract(country_input, stores_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Storedtgroups extraction
            storedtgroups_extract(country_input, dtgroups_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Storexcodegroups extraction
            storexcodegroups_extract(country_input, xcodegroups_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Lbatchstores extraction
            lbatchstores_extract(country_input, lbatchstores_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Store_group_stores extraction
            store_group_stores(country_input, store_group_stores_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Lbatches extraction
            lbatches_extract(country_input, lbatches_txt, excel_name_input, sheet_num_input, start_input, end_input)
            # Lbatchdtgroups extraction
            lbatchdtgroups_extract(country_input, lbatchdtgroups_txt, excel_name_input, sheet_num_input, start_input, end_input)
    except UnboundLocalError:
        sys.exit()

