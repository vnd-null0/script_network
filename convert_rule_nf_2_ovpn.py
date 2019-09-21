#from export_excel_hitcount_srx import create_dict_rule_srx, get_info_vlan
import re
import os
from netaddr import IPAddress, IPNetwork 
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import xlutils
from xlutils.margins import number_of_good_rows
from xlutils.view import Row, Col

def create_dict_rule_srx(list_conf_security_policy,_dict_vlan_):
    _dict_ = {}
    config_dict = {}
    protocol = []
    app = []
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
            #print(line)
            app.append(line)
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
            instance = line[line.find('groups') + 7:line.find('security') - 1]
            f_zone = line[line.find('from-zone') + 10:line.find('to-zone') - 1]
            t_zone = line[line.find('to-zone') + 8:line.find('policy') - 1]
            dPort = [x + '\n' for x in dPort]
            
            try:
                if sIP[0] != 'any\n' and IPNetwork(sIP[0]) in IPNetwork('10.79.0.0/16'):
                    #print(sIP[0])
                    tag = 'rule_ovpn'
                else:
                    tag = ''
            except:
                pass

            try:
                for route_ in _dict_vlan_:
                    if IPNetwork(dIP[0]) in IPNetwork(route_):
                        _vlan_ = _dict_vlan_[route_]['vlan_id']
                        dVLAN.append(_vlan_)
                        #print(str(dVLAN))
                        break
            except:
                pass

            dict_policy = {"name": name_policy, "instance": instance, "sourceIP": sorted(sIP), "destIP": sorted(dIP), \
                        "destVLAN": dVLAN, "destport": sorted(dPort), "protocol": sorted(protocol), "application": app, \
                        "f_zone": f_zone, "t_zone": t_zone, "tag": tag}
            
            _dict_.update([(name_policy, dict_policy)])
            config_dict.update([(name_policy, _config_)])

            dict_policy = {}
            sIP = []
            dIP = []
            app = []
            dVLAN = []
            list_dport = []
            dPort = []
            #sPort = []
            protocol = []
            _config_ = []
    
    #open(r'D:\PROGRAMING\_code_4_job_\info_dict_rule_ovpn.txt','w').write(str(_dict_))

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

def read_2_excel(lst_cfg):
    dict_vlan = get_info_vlan(lst_cfg)
    lst_security_policy = []
    rule_convert = []
    rule_ovpn_delete = []

    for file in lst_cfg:
        with open(file, 'r') as f:
            for line in f:
                if 'security policies' in line:
                    lst_security_policy.append(line)

    dict_rule, dict_cfg_ovpn = create_dict_rule_srx(lst_security_policy,dict_vlan)

    f1 = xlwt.Workbook()

    color = xlwt.easyxf(
        'align:wrap yes, vert centre,horiz centre;''pattern: pattern solid, fore_colour yellow;''font: colour red, name Arial, bold True;''borders: top double, bottom thin, left thin, right thin')
    #color1 = xlwt.easyxf('align:vert top,horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
    color2 = xlwt.easyxf(
        'align: wrap yes,vert top, horiz left;''borders:top dashed,right thin,left thin,bottom dashed;')
    
    f2 = f1.add_sheet('data_policy')
    for k in range(0, 14, 1):
        f2.col(k).width = 5000
    f2.row(0).height_mismatch = True
    f2.row(0).height = 500

    f2.write(0, 0, 'SRX5600_Newfarm', color)
    f2.write(0, 12, 'SRX-OVPN', color)
    f2.write(1, 0, 'policy_name', color)
    f2.write(1, 1, 'Instance', color)
    f2.write(1, 2, 'Action', color)
    f2.write(1, 3, 'Protocol', color)
    f2.write(1, 4, 'SourceIP', color)
    f2.write(1, 5, 'DestinationIP', color)
    f2.write(1, 6, 'DestinationPort', color)
    f2.write(1, 7, 'From_Zone', color)
    f2.write(1, 8, 'To_Zone', color)
    f2.write(1, 9, 'Application', color)
    f2.write(1, 10, 'Command_Rollback', color)
    f2.write(1, 11, 'Tag', color)
    f2.write(1, 12, 'Rule_convert', color)
    f2.write(1, 13, 'Command_Delete_on_SRX5600-NF', color)

    i = 2
    for key in dict_rule:
        f2.write(i, 0, dict_rule[key]['name'], color2)
        f2.write(i, 1, dict_rule[key]['instance'], color2)
        f2.write(i, 2, 'permit', color2)
        f2.write(i, 3, dict_rule[key]['protocol'], color2)
        f2.write(i, 4, dict_rule[key]['sourceIP'], color2)
        f2.write(i, 5, dict_rule[key]['destIP'], color2)
        f2.write(i, 6, dict_rule[key]['destport'], color2)
        f2.write(i, 7, dict_rule[key]['f_zone'], color2)
        f2.write(i, 8, dict_rule[key]['t_zone'], color2)
        f2.write(i, 9, dict_rule[key]['application'], color2)
        f2.write(i, 10, dict_cfg_ovpn[key], color2)
        f2.write(i, 11, dict_rule[key]['tag'], color2)
        
        if dict_rule[key]["tag"] == "rule_ovpn":
            for item in dict_rule[key]['sourceIP']:
                rule_convert.append('set security address-book global address '+str(item).replace("\n","")+" "+str(item))
                rule_ovpn_delete.append('delete security address-book global address '+str(item))
            for item in dict_rule[key]['destIP']:
                rule_convert.append('set security address-book global address '+str(item).replace("\n","")+" "+ str(item))
            for item in dict_rule[key]['sourceIP']:
                rule_convert.append('set security policies from-zone VLAN709 to-zone EXTERNAL policy '+\
                    str(dict_rule[key]['name'])+' match source-address '+str(item))
            for item in dict_rule[key]['destIP']:
                rule_convert.append('set security policies from-zone VLAN709 to-zone EXTERNAL policy '+\
                    str(dict_rule[key]['name'])+' match destination-address '+str(item))
            for item in dict_rule[key]['application']:
                rule_convert.append('set security policies from-zone VLAN709 to-zone EXTERNAL policy '+\
                    str(dict_rule[key]['name'])+' match application '+str(item))
            rule_convert.append('set security policies from-zone VLAN709 to-zone EXTERNAL policy '+\
                    str(dict_rule[key]['name'])+' then permit'+'\n')
            rule_ovpn_delete.append('delete groups '+dict_rule[key]['instance']+' security policies from-zone '+\
                    dict_rule[key]['f_zone']+' to-zone '+dict_rule[key]['t_zone']+' policy '+dict_rule[key]['name'])
            #print(rule_convert)
            #cmd_delete = 'delete security policies from-zone '+dict_rule[key]['f_zone']+' to-zone '+dict_rule[key]['t_zone']+' policy '+dict_rule[key]['name']

            f2.write(i, 12, rule_convert, color2)
            f2.write(i, 13, rule_ovpn_delete, color2)

        else:
            f2.write(i, 12, '', color2)
        
        rule_convert = []
        rule_ovpn_delete = []

        i +=1

    f1.save('move_rule_ovpn_nf_2_srx_ovpn.xls')

    return str(os.path.abspath('move_rule_ovpn_nf_2_srx_ovpn.xls'))

if __name__ == '__main__':
    list_config = ['srx5600_N0_config_backup_09072019.txt']
    read_2_excel(list_config)