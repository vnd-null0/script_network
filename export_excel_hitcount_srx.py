import re
import os
from netaddr import *
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import xlutils
from xlutils.margins import number_of_good_rows
from xlutils.view import Row, Col

def read_tsv(file_hitcount):
    list_dict_hitcount = []

    with open(file_hitcount, 'r') as fh:
        key = ['index', 'f_zone', 't_zone', 'policy_name', 'p_count']
        length = len(key)
        i = 0
        for line in fh.readlines():
            i += 1
            if i == 1:
                continue
            else:
                value = line.split()
                dic = {key[x]: value[x] for x in range(length)}
                list_dict_hitcount.append(dic)
    fh.close()

    return list_dict_hitcount

'''
def filter_list_rule_srx(file_config_srx):
    list_rule_srx = []

    with open(file_config_srx, 'r') as f:
        for line in f:
            if 'security policies' in line:
                list_rule_srx.append(line)
    f.close()

    return list_rule_srx
'''

def create_dict_rule_srx(list_conf_security_policy,_dict_vlan_):
    _dict_ = {}
    config_dict = {}
    protocol = []
    sIP = []
    dIP = []
    dPort = []
    #sPort = []
    _config_ = []
    dVLAN = []
    #hostname = 'NO_IDENTIFY'

    for line in list_conf_security_policy:
        if 'source-address' in line:
            _config_.append(line)
            line = line[line.find('source-address') + 15:]
            sIP.append(line)
        elif 'destination-address' in line:
            _config_.append(line)
            line = line[line.find('destination-address') + 20:]
            dIP.append(line)
        elif 'application' in line:
            _config_.append(line)
            line = line[line.find('application') + 12:]
            if 'any' not in line:
                if 'tcp' in line and 'udp' in line:
                    protocol = ['tcp\n', 'udp\n']
                    list_dport = re.findall('[\d\-]+', line)
                    dPort.extend(list_dport)
                elif 'tcp' in line:
                    protocol = ['tcp\n']
                    list_dport = re.findall('[\d\-]+', line)
                    dPort.extend(list_dport)
                elif 'udp' in line:
                    protocol = ['udp\n']
                    list_dport = re.findall('[\d\-]+', line)
                    dPort.extend(list_dport)
                else:
                    protocol = line
                    dPort.append(line)
            else:
                if 'tcp' in line and 'udp' in line:
                    protocol = ['tcp\n', 'udp\n']
                    dPort.append('any')
                elif 'tcp' in line:
                    protocol = ['tcp\n']
                    dPort.append('any')
                elif 'udp' in line:
                    protocol = ['udp\n']
                    dPort.append('any')
                else:
                    protocol = ['any\n']
                    dPort.append('any')
            try:
                dPort = [x for x in dPort if x != '-']
            except:
                pass
        elif 'protocol' in line:
            _config_.append(line)
            protocol.append(line[line.find('protocol') + 9:])
        elif 'then permit' in line:
            _config_.append(line)
            name_policy = line[line.find('policy') + 7:line.find('then') - 1]
            dPort = [x + '\n' for x in dPort]

            try:
                for route_ in _dict_vlan_:
                    if IPNetwork(dIP[0]) in IPNetwork(route_):
                        _vlan_ = _dict_vlan_[route_]['vlan_id']
                        dVLAN.append(_vlan_)
                        #print(str(dVLAN))
                        break
            except:
                pass

            dict_policy = {"name": name_policy, "sourceIP": sorted(sIP), "destIP": sorted(dIP), "destVLAN": dVLAN,\
                            "destport": sorted(dPort), "protocol": sorted(protocol)}
            
            _dict_.update([(name_policy, dict_policy)])
            config_dict.update([(name_policy, _config_)])

            dict_policy = {}
            sIP = []
            dIP = []
            dVLAN = []
            list_dport = []
            dPort = []
            #sPort = []
            protocol = []
            _config_ = []
    
    open(r'D:\PROGRAMING\_code_4_job_\info_dict_rule_ovpn.txt','w').write(str(_dict_))

    return _dict_, config_dict

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
                info_dict_vlan.update([(address,{'hostname':hostname,'vlan_id':vlan_id})])
    
    return info_dict_vlan

def read_2_excel(lst_config,file_hitcount):
    _list_of_dict_ = read_tsv(file_hitcount)
    _dict_vlan_ = get_info_vlan(lst_config)
    list_security_policy_ovpn = []

    for file in lst_config:
        with open(file, 'r') as f:
            for line in f:
                if 'security policies from-zone VLAN709 to-zone EXTERNAL' in line:
                    list_security_policy_ovpn.append(line)

    dict_srx_ovpn,dict_config_ovpn = create_dict_rule_srx(list_security_policy_ovpn,_dict_vlan_)

    f1 = xlwt.Workbook()

    color = xlwt.easyxf(
        'align:wrap yes, vert centre,horiz centre;''pattern: pattern solid, fore_colour yellow;''font: colour red,name Arial, bold True;''borders: top double, bottom thin, left thin, right thin')
    color1 = xlwt.easyxf('align:vert top,horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
    color2 = xlwt.easyxf(
        'align: wrap yes,vert top, horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
        
    f2 = f1.add_sheet('policy_hitcount')
    f3 = f1.add_sheet('data_policy')
    for i in range(1, 6, 1):
        f2.col(i).width = 6000
        #f3.col(i).width = 5000
    for k in range(0, 11, 1):
        #f2.col(i).width = 5000
        f3.col(k).width = 5000    
    f2.col(0).width = 3000
    f2.row(0).height_mismatch = True
    f2.row(0).height = 500
    #f3.col(0).width = 3000
    f3.row(0).height_mismatch = True
    f3.row(0).height = 500

    f2.write(0, 0, 'SRX_OVPN', color)
    f2.write(1, 0, 'Index', color)
    f2.write(1, 1, 'From_Zone', color)
    f2.write(1, 2, 'To_Zone', color)
    f2.write(1, 3, 'policy_name', color)
    f2.write(1, 4, 'hitcount', color)
    f2.write(1, 5, 'Note', color)

    f3.write(0, 0, 'SRX_OVPN', color)
    f3.write(1, 0, 'policy_name', color)
    f3.write(1, 1, 'Action', color)
    f3.write(1, 2, 'Protocol', color)
    f3.write(1, 3, 'SourceIP', color)
    #f3.write(1, 4, 'SourcePort', color)
    f3.write(1, 4, 'DestinationIP', color)
    f3.write(1, 5, 'DestinationPort', color)
    f3.write(1, 6, 'Destination_VLAN', color)
    f3.write(1, 7, 'Command_Rollback', color)
    f3.write(1, 8, 'From_Zone', color)
    f3.write(1, 9, 'To_Zone', color)
    f3.write(1, 10, 'hitcount', color)
    f3.write(1, 11, 'Delete', color)

    
    x = 2
    for key in range(len(_list_of_dict_)):
        f2.write(x, 0, _list_of_dict_[key]['index'], color2)
        f2.write(x, 1, _list_of_dict_[key]['f_zone'], color2)
        f2.write(x, 2, _list_of_dict_[key]['t_zone'], color2)
        f2.write(x, 3, _list_of_dict_[key]['policy_name'], color2)
        f2.write(x, 4, _list_of_dict_[key]['p_count'], color2)
        f2.write(x, 5, 'Delete' if _list_of_dict_[key]['p_count'] == "0" else 'Keep', color2)

        x += 1
    
    y = 2
    for key in dict_srx_ovpn:
        f3.write(y, 0, dict_srx_ovpn[key]['name'], color1)
        f3.write(y, 1, 'permit', color1)
        f3.write(y, 2, dict_srx_ovpn[key]['protocol'], color2)
        f3.write(y, 3, dict_srx_ovpn[key]['sourceIP'], color2)
        #f3.write(y, 4, dict_srx_ovpn[key]['sourceport'], color2)
        f3.write(y, 4, dict_srx_ovpn[key]['destIP'], color2)
        f3.write(y, 5, dict_srx_ovpn[key]['destport'], color2)
        f3.write(y, 6, dict_srx_ovpn[key]['destVLAN'], color2)
        f3.write(y, 7, dict_config_ovpn[key], color1)
        f3.write(y, 8, 'VLAN709', color2)
        f3.write(y, 9, 'EXTERNAL', color2)
        #f3.write(y, 11, 'Note', color2)

        try:
            for i in range(len(_list_of_dict_)):
                if  dict_srx_ovpn[key]['name'] == _list_of_dict_[i]['policy_name']:
                    cmd_delete = 'delete security policies from-zone VLAN709 to-zone EXTERNAL policy ' + dict_srx_ovpn[key]['name']
                    f3.write(y, 10, _list_of_dict_[i]['p_count'], color2)
                    f3.write(y, 11,  cmd_delete if _list_of_dict_[i]['p_count'] == '0' else '', color2)
                    break
        except:
            pass
        
        y +=1

    f1.save('hitcount_ovpn.xls')

    return str(os.path.abspath('hitcount_ovpn.xls'))

if __name__ == '__main__':
    list_config = ['config_backup_srx_ovpn_15082019.txt','srx5800_config_backup_06052019.txt','ex9214_01_config_backup.txt',\
                    'ex9214_02_config_backup.txt','srx5600_N0_config_backup_09072019.txt','srx3600_dr_config_backup_06052019.txt',\
                    'EX4200VC_FV_config_backup.txt','srx123pay_config_backup_06052019.txt']
    read_2_excel(list_config,'hitcount_ovpn.txt')