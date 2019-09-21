import re
import os
from netaddr import *
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import xlutils
from xlutils.margins import number_of_good_rows
from xlutils.view import Row, Col

def get_info_vlan(list_config_device): #trả về info vlan-id khi lookup IP, include SRX5800, EX9214
    info_dict_vlan = {}

    for file in list_config_device:
        hostname = 'NO_IDENTIFY'
        conf_lst = open(file, 'r').readlines()
        for line in conf_lst:
            try:
                hostname = re.search('set[a-zA-Z0-9\s]+system host-name (.*)_N[0-1]', line).group(1)
                break
            except:
                pass
            try:
                hostname = re.search('set[a-zA-Z0-9\s]+system host-name (.*)', line).group(1)
                break
            except:
                pass

        for line in conf_lst:
            if re.search('set(.*)interfaces (.*) unit (.*) family inet address (.*)',
                         line) and 'interfaces fxp' not in line and 'interfaces lo0' not in line:
                key = re.search('set(.*)interfaces (.*) unit (.*) family inet address (.*)', line)
                vlan_id = key.group(3)
                _address_ = key.group(4)
                address = IPNetwork(_address_[0:_address_.find('/')+3]).cidr
                info_dict_vlan.update([(vlan_id,{'hostname':hostname,'ip_range':str(address),'gateway':_address_})])
    
    return info_dict_vlan

def parse_2_excel(lst_cfg):
	dict_info_vlan = get_info_vlan(lst_cfg)
	
	f1 = xlwt.Workbook()
	
	theme_1 = xlwt.easyxf(
		'align:wrap yes, vert centre,horiz centre;''pattern: pattern solid, fore_colour yellow;''font: colour red,name Arial, bold True;''borders: top double, bottom thin, left thin, right thin')
	theme_2 = xlwt.easyxf('align:vert top,horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
	theme_3 = xlwt.easyxf('align: wrap yes,vert top, horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
	
	f2 = f1.add_sheet('info_vlan_gateway')
	
	for i in range(2, 6, 1):
		f2.col(i).width = 6000
	
	f2.col(0).width = 3000
	f2.col(1).width = 7000
	f2.row(0).height_mismatch = True
	f2.row(0).height = 500
	f2.row(1).height_mismatch = True
	f2.row(1).height = 500

	f2.write(0, 0, 'INFO_VLAN_GATEWAYS', theme_1)
	f2.write(1, 0, 'INDEX', theme_1)
	f2.write(1, 1, 'DEVICE', theme_1)
	f2.write(1, 2, 'VLAN_ID', theme_1)
	f2.write(1, 3, 'IP_GATEWAY', theme_1)
	f2.write(1, 4, 'RANGE_IP', theme_1)

	x = 2
	m = 1
	for key in dict_info_vlan:
		f2.write(x, 0, m, theme_2)
		f2.write(x, 1, dict_info_vlan[key]['hostname'], theme_2)
		f2.write(x, 2, key, theme_2)
		f2.write(x, 3, dict_info_vlan[key]['gateway'], theme_2)
		f2.write(x, 4, dict_info_vlan[key]['ip_range'], theme_2)
		
		m += 1
		x += 1

	f1.save('parse_vlan_gateway.xls')

	return str(os.path.abspath('parse_vlan_gateway.xls'))

if __name__ == '__main__':
    list_config = ['mx480_01_config_backup.txt','config_srx_ovpn_backup.txt','srx5800_config_backup.txt','ex9214_01_config_backup.txt',\
                    'ex9214_02_config_backup.txt','srx5600_N0_config_backup.txt','srx3600_dr_config_backup.txt',\
                    'EX4200VC_FV_config_backup.txt']
					
    parse_2_excel(list_config)