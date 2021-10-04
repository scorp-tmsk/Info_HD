import wmi
import binascii


def creat_list_pc():
    # create list with name pc
    list_pc = []
    with open('name_pc.txt', encoding='utf-8') as pc:
        for line in pc:
            list_pc.append(line.strip())
    return list_pc


def out_info(info_pc):
    # Выводим информацию в файл из info_all_pc
    if len(info_pc) != 0:
        with open('info_pc_hd.rtf', 'w') as inf:
            for key in info_pc.keys():
               # inf.write('Name PC:\t' + key + '\n\n')
                for key1, value1 in info_pc[key].items():
                   inf.write(key + '\t' +  key1 + ':\t' + value1 + '\n\n')
                # inf.write('*' * 20 + '\n\n')


def out_info_error(error_conect_pc):
    # Выводим информацию в файл из bad_pc
    if len(error_conect_pc) != 0:
        with open('error_conect_pc_hd.rtf', 'w') as inf:
            for i in error_conect_pc:
                inf.write(str(i) + '\n')


def info_hd(name_pc):
    hd_info = {}
    print(name_pc)
    try:
        # "{}".format(name_pc)
        conect_pc = wmi.WMI("{}".format(name_pc))
        for info in conect_pc.Win32_diskdrive():
            if len(str(info.SerialNumber.strip())) >= 40:
                sn = str(binascii.a2b_hex(info.SerialNumber).strip())
                hd_info.update({info.model.strip(): sn
                                + '\t' + str(info.MediaType.strip()) + '\t'
                                + str(info.InterfaceType.strip())})
            else:
                hd_info.update({info.model.strip(): str(info.SerialNumber.strip())
                                + '\t' + str(info.MediaType.strip()) + '\t'
                                + str(info.InterfaceType.strip())})
        return(hd_info)
    except:
        return name_pc


error_conect_pc_hd = []
info_pc_hd = {}
# создает список имен компьютеров
list_name_pc = creat_list_pc()
# Проверка типа list or dict
for name in list_name_pc:
    result = info_hd(name)
    if type(result) == dict:
        info_pc_hd[name] = result
    else:
        error_conect_pc_hd.append(result)

out_info(info_pc_hd)

out_info_error(error_conect_pc_hd)
