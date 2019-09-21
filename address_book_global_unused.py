import re
import os
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import xlutils
from xlutils.margins import number_of_good_rows
from xlutils.view import Row, Col

def create_list_add_book_unused(conf_device):
    lst_address_book = []
    lst_address_policy = []

    for line in conf_device:
        if re.search('set security address-book global address (\d+\.\d+\.\d+\.\d+/\d+) (\d+\.\d+\.\d+\.\d+/\d+)', line):
            address_book = re.search('set security address-book global address (\d+\.\d+\.\d+\.\d+/\d+) (\d+\.\d+\.\d+\.\d+/\d+)', line).group(2)
            lst_address_book.append(address_book)
        elif re.search('set(.*)security policies (.*) match(.*)address (\d+\.\d+\.\d+\.\d+/\d+)', line):
            address_policy = re.search('set(.*)security policies (.*) match (.*)address (\d+\.\d+\.\d+\.\d+/\d+)', line).group(4)
            lst_address_policy.append(address_policy)
    
    #print(str(lst_address_policy))
    #print(str(lst_address_book))
    lst_add_unused = list(set(lst_address_book) ^ set(lst_address_policy))
    #print(str(lst_add_unused))

    return lst_add_unused

def create_dict_mapping_device(list_config):

    info_dict = {}
    lst_add_book = []
    for file in list_config:
        hostname = 'NO_IDENTIFY'
        cfg_lst = open(file, 'r').readlines()
        for line in cfg_lst:
            try:
                hostname = re.search('set[a-zA-Z0-9\s]+system host-name (.*)(_|-)N[0-1]', line).group(1)
                break
            except:
                pass
            try:
                hostname = re.search('set[a-zA-Z0-9\s]+system host-name (.*)', line).group(1)
                break
            except:
                pass

        lst_add_book = create_list_add_book_unused(cfg_lst)
        info_dict.update([(hostname,lst_add_book)])

        lst_add_book = []

    return info_dict

def parse_to_excel(lst_conf):
    dict_address_book_unused = create_dict_mapping_device(lst_conf)

    f1 = xlwt.Workbook()

    color = xlwt.easyxf(
        'align:wrap yes, vert centre,horiz centre;''pattern: pattern solid, fore_colour yellow;''font: colour red,name Arial, bold True;''borders: top double, bottom thin, left thin, right thin')
    color1 = xlwt.easyxf('align:vert top,horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
    color2 = xlwt.easyxf(
        'align: wrap yes,vert top, horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
    
    for key in dict_address_book_unused:
        worksheet = f1.add_sheet('%s' %key)

        worksheet.col(0).width = 5000
        worksheet.row(0).height_mismatch = True
        worksheet.row(0).height = 500

        worksheet.write(0, 0, key, color)
        worksheet.write(1, 0, 'Address_Unused', color)
        worksheet.write(1, 1, 'Command', color)

        x = 2
        for add in range(len(dict_address_book_unused[key])):
            worksheet.write(x, 0, dict_address_book_unused[key][add], color2)
            cmd_delete = 'delete security address-book global address ' + dict_address_book_unused[key][add]
            worksheet.write(x, 1, cmd_delete, color2)

            x += 1

    f1.save('address_book_global_unused_srx.xls')

    return str(os.path.abspath('address_book_global_unused_srx.xls'))

if __name__ == '__main__':
    list_cfg = ['config_backup_srx_ovpn_06052019.txt','srx5800_config_backup_06052019.txt','srx5600_N0_config_backup_06052019.txt',\
                'srx3600_dr_config_backup_06052019.txt','srx123pay_config_backup_06052019.txt']
    parse_to_excel(list_cfg)