#SteamsInfo modified to StreamInfo on 27th Feb
#Modified parameter Stream.TransmissionModeValue to provide type and value in type:value format
#5G UL is classified based on ether type 0x8951 and DL is based on 1019 or 2019 in teid
#VLAN tag added to the interface based on vlanEnable flag
#ARP type header is handled
#VPorts for connected and unconnected interface added, Vlan variable name is modified according to FW mapper file
#Connected via interface name accessed from 2 (sequence number)
#Adding handling of auto parameter
#17thJune: Parameters index corrected in createTemplateType,handling of auto is modified to provide values in cases its non zero,so
#          template filling would be easier

import json, os, sys
import pandas as pd
from pandas import DataFrame
from pandas import ExcelWriter
import xlrd
import copy


template_field_list = [
'ethernet.header.destinationAddress-1',
'ethernet.header.etherType-3',
'vlan.header.protocolID-4',
'ipv4.header.protocol-25',
'ipv6.header.nextHeader-5',
'udp.header.dstPort-2',
'GtpF15g.header.GtpF15gTEID-9',
'gtpf1.header.teid-9',
'gtpu.header.teid-9'
]
dst_mac_index = 23
eth_type_index = 24
vlan_type_index = 27
ipv4_proto_index = 44
ipv6_proto_index = 52
udp_dstport_index = 69
gtpf15_teid_index = 123
gtpf1_teid_index = 158
gtpu_teid_index = 104

def createTemplateType(dict):
    global hw_type
    global rel_type
    #check MAC address list
    ul_lastByte = ["ff:ff:ff:ff:ff:02","ff:ff:ff:ff:ff:07","ff:ff:ff:ff:ff:21","ff:ff:ff:ff:ff:22","02:40:43:80:10:08","02:40:43:80:20:08"]
    ip = "v4"
    traffic_dir=""
    rat = ''
    traffic = ''
    vlan_en = False
    up_traffic = False
    try:
        if dict[dst_mac_index] in ul_lastByte:
            traffic_dir = "UL"
        else:
            traffic_dir = "DL"
    except KeyError as e:
        print ("error in decoding MAC address in ethernet header for stream ",dict[3])
    #eth type check
    try:
        if dict[eth_type_index] == "0x806":
            rat,traffic = "Generic","ARP"
        elif dict[eth_type_index] == "0x800":
            ip = "v4"
        elif dict[eth_type_index] == "0x86dd":
            ip = "v6"
        elif dict[eth_type_index] == "0x8951":
            traffic_dir = "UL"
        elif dict[eth_type_index] == "0x8100":
	        vlan_en = True
        else:
            traffic_dir = "DL"
    except KeyError as e:
        print ("error in decoding Eth type for stream", dict[3])

    #vlan type check
    try:
        if vlan_en:
            if dict[vlan_type_index] == "0x806":
                rat,traffic =  "Generic","ARP"   
            elif dict[vlan_type_index] == "0x800":
                ip = "v4"
            elif dict[vlan_type_index] == "0x86dd":
                ip = "v6"
            elif dict[vlan_type_index] == "0x8951":
                traffic_dir = "UL"
                #if "SRAN" in rel_type.upper():
				#    rat = "LTE"
            else:
                traffic_dir = "DL"
    except KeyError as e:
        print ("error in decoding VLAN Type for stream", dict[3])
    
    try:
        if ip == "v4" and dict[ipv4_proto_index] == "17":
            up_traffic = True
        elif ip == "v6" and dict[ipv6_proto_index] == "17":
            up_traffic = True
        elif ip == "v6" and dict[ipv6_proto_index] == "44":
            rat,traffic = "Generic","IPv6_Fragmentation"
        else:
            up_traffic = False
            rat,traffic = "Generic","IP"
    except KeyError as e:
        print ("error in decoding Protocol Type for stream", dict[3])
    try:
        if up_traffic:
            if dict[udp_dstport_index] > "49000":
                rat = "Iub"
            elif dict[udp_dstport_index] == "52152":
                rat = "C1"
            elif dict[udp_dstport_index] == "2152":
                if dict[gtpf15_teid_index] or dict[gtpf1_teid_index]:
                    rat = "F1"
                    print ("Its F1",dict[3])
                elif dict[gtpu_teid_index]:
                    rat = "S1"
                    print ("Its S1",dict[3])
            else:
                rat,traffic = "Generic","IP"

    except KeyError as e:
        print ("error in decoding UDP Port/Teid for stream", dict[3])
    if rat == '':
        if "CBTS" in rel_type.upper():
            rat = "C1"
        if "5G" in rel_type.upper():
            rat = "F1"
        if "SRAN" in rel_type.upper():
            rat = "S1"
    if rat == "Generic":
        template_type = '_'.join([rat, traffic, 'Template'])
    elif traffic_dir == "UL":
        template_type = '_'.join([rat, traffic_dir,hw_type, 'Template'])
    elif traffic_dir == "DL" and rat == "C1":
        template_type = '_'.join([rat, traffic_dir,hw_type, 'Template'])
    elif traffic_dir == "DL":
        template_type = '_'.join([rat, traffic_dir, 'Template'])
    else:
	    template_type = "Not Defined"
    return template_type

def createInterfaceDict(values, dataframe_dict):
    global file_path
    with open(file_path,"r") as f:
        file_read = f.read()
    json_file = json.loads(file_read)
    vports = json_file["vport"]
    dictionaryList = {'my_dictionary':[]}
    dictionary = {}
    vportList = []
    for vport in vports:
        for interface in vport["interface"]:
            if interface["unconnected"]["connectedVia"] == None:
                dictionary = {}
                dictionary[dataframe_dict['SequenceNumber']['InterfaceName']] = interface["description"]
                dictionary[dataframe_dict['SequenceNumber']['VportName']] = interface["xpath"].split('/')[1]
                if "ipv4" in interface.keys():            
                    dictionary[dataframe_dict['SequenceNumber']['IPv4address']] = interface["ipv4"]["ip"]
                    dictionary[dataframe_dict['SequenceNumber']['IPv4Subnet']] = interface["ipv4"]["maskWidth"]
                    dictionary[dataframe_dict['SequenceNumber']['IPv4Gateway']] = interface["ipv4"]["gateway"]
                if "ipv6" in interface.keys():
                    for ip in interface["ipv6"]:
                        dictionary[dataframe_dict['SequenceNumber']['IPv6address']] = ip["ip"]
                        dictionary[dataframe_dict['SequenceNumber']['IPv6Subnet']] = ip["prefixLength"]
                        dictionary[dataframe_dict['SequenceNumber']['IPv6Gateway']] = ip["gateway"]
                dictionary[dataframe_dict['SequenceNumber']['MTU']] = interface["mtu"]
                if interface["vlan"]["vlanEnable"]:
                    dictionary[dataframe_dict['SequenceNumber']['Vlans']] = interface["vlan"]["vlanId"]
                dictionaryList['my_dictionary'].append(dictionary)
            else:
                dictionary = {}
                dictionary[dataframe_dict['SequenceNumber']['UI_InterfaceName']] = interface["description"]
                dictionary[dataframe_dict['SequenceNumber']['UI_VportName']] = interface["xpath"].split('/')[1]
                if "ipv4" in interface.keys():            
                    dictionary[dataframe_dict['SequenceNumber']['UI_IPv4address']] = interface["ipv4"]["ip"]
                    dictionary[dataframe_dict['SequenceNumber']['UI_IPv4Subnet']] = interface["ipv4"]["maskWidth"]
                if "ipv6" in interface.keys():
                    for ip in interface["ipv6"]:
                        dictionary[dataframe_dict['SequenceNumber']['UI_IPv6address']] = ip["ip"]
                        dictionary[dataframe_dict['SequenceNumber']['UI_IPv6Subnet']] = ip["prefixLength"]
                connectedVia = interface["unconnected"]["connectedVia"] #/vport[1]/interface[3]
                connectedVia = int(connectedVia.split('[')[-1][:-1])
#               2- sequence number for interface name
                if len(dictionaryList['my_dictionary']) >= connectedVia:
                    print ("connected via interface:",dictionaryList['my_dictionary'][connectedVia-1][2])
                    connectedVia = dictionaryList['my_dictionary'][connectedVia-1][2]
                else:
                    print ("connectedVia interface index",connectedVia)
                    print ("dict len",len((dictionaryList['my_dictionary'])))
                    print ("error in handling interface index and is due to JSON file error")
                dictionary[dataframe_dict['SequenceNumber']['UI_ConnectedVia']] = connectedVia
                if interface["vlan"]["vlanEnable"]:
                    dictionary[dataframe_dict['SequenceNumber']['UI_Vlans']] = interface["vlan"]["vlanId"]
                dictionaryList['my_dictionary'].append(dictionary)

        vportList.append(vport["xpath"].split('/')[1])
    return dictionaryList, vportList

def createInterfaceSheet(excelFile):
    global writer
    dataframe = pd.read_excel(excelFile, sheet_name = 'GlobalInfo')
    values = dataframe['SequenceNumber'].values
    dataframe_dict = pd.read_excel(excelFile, sheet_name = 'GlobalInfo', index_col=1).to_dict()
    dictList, vportList = createInterfaceDict(values, dataframe_dict)
    df = pd.DataFrame(dictList)
    df2 = df['my_dictionary'].apply(pd.Series)
    df3 = df2.transpose()
    df3.columns = range(1, len(df3.columns)+1)
    df3.index.name = "SequenceNumber"
    merged = pd.merge(left=df3, left_index=True, right=dataframe, right_on="SequenceNumber", how='right')
    merge = merged.set_index("SequenceNumber")
    cols = merge.columns.tolist()
    cols = cols[-2:] + cols[:-2]
    merge = merge[cols]
	#testing for getting port index
    #port_index = merged["NokiaVariables"]
    #port_index1 = port_index.columns.tolist()
    #print ("**********port index",merged.iloc[3])
    merge.at[28, 'DefaultValues'] = ','.join(vportList)
    merge.to_excel(writer,'GlobalInfo')
    #df3.to_excel(writer,'Sheet2')

def createDictionary(values, dataframe_dict):
    global file_path
    with open(file_path,"r") as f:
        file_read = f.read()
    json_file = json.loads(file_read)
    streams = json_file["traffic"]["trafficItem"]
    dictList = {'my_dict':[]}
    for i in streams:
        #print "Stream Name: ", i["name"]
        dict = {}
        dict[dataframe_dict['SequenceNumber']['Stream Name']] = i["name"]
        for j in i["configElement"]:
            dict[dataframe_dict['SequenceNumber']['Frame Size Type']] = j["frameSize"]["type"]
            if j["frameSize"]["type"] == "fixed":
                dict[dataframe_dict['SequenceNumber']['Frame Size Fixed']] = j["frameSize"]["fixedSize"]
            elif j["frameSize"]["type"] == "increment":
                dict[dataframe_dict['SequenceNumber']['Frame Size Increment']] = "[From:"+j["frameSize"]["incrementFrom"]+",Step:"+j["frameSize"]["incrementStep"]+",To:"+j["frameSize"]["incrementTo"]+"]"
            elif j["frameSize"]["type"] == "random":
                dict[dataframe_dict['SequenceNumber']['Frame Size Random']] = "[Min:"+str(j["frameSize"]["randomMin"])+",Max:"+str(j["frameSize"]["randomMax"])+"]"
            elif j["frameSize"]["type"] == "presetDistribution":
                dict[dataframe_dict['SequenceNumber']['Frame Size imix']] = j["frameSize"]["presetDistribution"]
            elif j["frameSize"]["type"] == "weightedPairs":
                dict[dataframe_dict['SequenceNumber']['Frame Size Weighted Pairs']] = j["frameSize"]["weightedPairs"]
            elif j["frameSize"]["type"] == "quadGaussian":
                dict[dataframe_dict['SequenceNumber']['Frame Size Quad Gaussian']] = j["frameSize"]["quadGaussian"]
            dict[dataframe_dict['SequenceNumber']['Rate Type']] = j["frameRate"]["type"]
            if dict[dataframe_dict['SequenceNumber']['Rate Type']] == "bitsPerSecond":
                dict[dataframe_dict['SequenceNumber']['Rate Value']] = str(j["frameRate"]["rate"]) +" "+ j["frameRate"]["bitRateUnitsType"]
            else:
                dict[dataframe_dict['SequenceNumber']['Rate Value']] = j["frameRate"]["rate"]
            dict[dataframe_dict['SequenceNumber']['Payload Type']] = j["framePayload"]["type"]
            dict[dataframe_dict['SequenceNumber']['Transmission Mode Type']] = j["transmissionControl"]["type"]
            if j["transmissionControl"]["type"] == "fixedFrameCount":
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "frameCount:" + str (j["transmissionControl"]["frameCount"])
            elif j["transmissionControl"]["type"] == "fixedIterationCount":
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "iterationCount:" + str(j["transmissionControl"]["iterationCount"])
            elif j["transmissionControl"]["type"] == "fixedDuration":
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "duration:" + str(j["transmissionControl"]["duration"])
            elif j["transmissionControl"]["type"] == "continuous":
                #print ("continous stream")
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "duration:" + str(j["transmissionControl"]["duration"])
            elif j["transmissionControl"]["type"] == "auto":
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "minGapBytes:" + str(j["transmissionControl"]["minGapBytes"])
            elif j["transmissionControl"]["type"] == "custom":
                dict[dataframe_dict['SequenceNumber']['Transmission Mode Value']] = "[Burst:"+str(j["transmissionControl"]["burstPacketCount"])+"," + "Gap:"+str(j["transmissionControl"]["minGapBytes"])+"]"
            for k in j["stack"]:
                for l in k["field"]:
                    #print "xpath: ", l["xpath"].split("=")[-1][:-1], "\t singleValue: ", l["singleValue"]
                    x = l["xpath"].split("=")[-1][:-1]
                    x=x.strip()
                    #dict [x] = l["singleValue"]
                    if x in values:
                        if l["valueType"]=="singleValue":
                            #if l["auto"]:
                            #    val = "auto"                                 
                            #if l["singleValue"] != "0":
                            #    val = l["singleValue"]
                            if l["auto"] and l["xpath"].split("=")[-1].split("']")[0].strip().split("'")[-1] not in template_field_list:
                                val = "auto"
                            else:
                                val = l["singleValue"]
                        elif l["valueType"]=="decrement":
                            val = "Decrement[start:"+l["startValue"]+",step:"+l["stepValue"]+",count:"+l["countValue"]+"]"
                        elif l["valueType"]=="increment":
                            val = "Increment[start:"+l["startValue"]+",step:"+l["stepValue"]+",count:"+l["countValue"]+"]"
                        elif l["valueType"]=="valueList":
                            val = "List["+','.join(l["valueList"])+']'

                        if 'ipv4.header.flags.fragment-21' in x and l["fieldValue"] == "May fragment":
                            dict[dataframe_dict['SequenceNumber']['MF']] = val
                            continue
                        elif 'ipv4.header.flags.fragment-21' in x and l["fieldValue"] != "May fragment":
                            dict[dataframe_dict['SequenceNumber']['DF']] = val
                            continue
                        #IPV6 fragmentation header handling
                        if 'ipv6.header.nextHeader-5' in x and val == "44":
                            print ("IPv6 fragmentation header exist")

                        if 'header.etherType-3' in x or 'vlan.header.protocolID-4' in x :
                            if not val == "auto":
                                val = str("0x") + str(val)

                        dict[dataframe_dict['SequenceNumber'][x]] = val
        for j in i["endpointSet"]:
            for x in j["destinations"]:
                dict[dataframe_dict['SequenceNumber']['Src End Point']] = x.split('/')[1]
            for x in j["sources"]:
                dict[dataframe_dict['SequenceNumber']['Dest End Point']] = x.split('/')[1]
        for j in i["tracking"]:
            dict[dataframe_dict['SequenceNumber']['Traffic Item']] = "enabled" if j["trackBy"] != [] else "disabled"
        for j in i["highLevelStream"]:
            dict[dataframe_dict['SequenceNumber']['CRC']] = j["crc"]
        ret = createTemplateType(dict)        
        #print ("returned template type:",ret)
        dict[dataframe_dict['SequenceNumber']['Template Type']] = ret
        dictList['my_dict'].append(dict)

    return dictList
   
def createExcel(excelFile):
    new_list = []
    final_list_appended = []
    global writer
    dataframe = pd.read_excel(excelFile, sheet_name = 'MappingInfo')
    values = dataframe['IxiaVariables'].values
    dataframe_dict = pd.read_excel(excelFile, sheet_name = 'MappingInfo', index_col=1).to_dict()
    values = values.tolist()
    dictList_original = createDictionary(values, dataframe_dict)
    dictList = copy.deepcopy(dictList_original)
    
    
    print (len(dictList_original['my_dict']))
    #dictList = createDictionary(values, dataframe_dict)
    
    initial_list_inside_dict = dictList['my_dict']
    print (len(initial_list_inside_dict))
    for l in initial_list_inside_dict:
        
        if 'LTE' in l[3].split('_') and 'S1' in l[4].split('_'):
            l[3] = l[3].replace("LTE","5G")
            l[4] = l[4].replace("S1","F1")
            l[122] = l[110]
            l[120] = l[108]
            l[123] = l[104]
            l[116] = l[105]
            l[118] = l[106]
            l[119] = l[107]
            l[121] = l[109]
            
            
            
            new_list.append(l)
        #else:
            #print (l[3] + ' no' + l[4])
    
    print ("original list is *****************************\n**************\n\n")
    print (dictList_original['my_dict'])
    final_list_appended = dictList_original['my_dict'] + new_list
    print (final_list_appended)
    print ("***************************\n***********************")
    dictList = {'my_dict':final_list_appended}
    #final_list_appended = dictList
    print (dictList)            
    
    
    df = pd.DataFrame(dictList)
    df2 = df['my_dict'].apply(pd.Series)
    df3 = df2.transpose()
    df3.columns = range(1, len(df3.columns)+1)
    df3.index.name = "SequenceNumber"
    merged = pd.merge(left=df3, left_index=True, right=dataframe, right_on="SequenceNumber", how='right')
    merge = merged.set_index("SequenceNumber")
    cols = merge.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    merge = merge[cols]
    merge = merge.rename(columns={'NokiaVariables': 'StreamParameters'})
    merge = merge.drop(['IxiaVariables'], axis=1)

    merge.to_excel(writer,'StreamInfo')
    #df3.to_excel(writer,'Sheet2')
    
writer = None
file_path = None
def main(rootDir):
    global writer
    global file_path
    excelPath = '\Config\Setup_Config\SCT\Ixia'
    directorylist = []
    for dirs in os.listdir(rootDir):
        for root, dir, files in os.walk(os.path.join(rootDir, dirs) + excelPath):
            directorylist.append(root)
    for i in range(len(directorylist)):
        for files in os.listdir(directorylist[i]):
            if(files.endswith(".json")):
                jsonfile = os.path.join(directorylist[i], files)
                file_path = jsonfile
                outputFilePath = file_path.split('.json')[0] + '.xlsx'
                writer = ExcelWriter(outputFilePath)
                createExcel(excelFile)
                createInterfaceSheet(excelFile) 
                writer.save()

#excelFile = input('Path to Input Excel File: ')
excelFile = "C:\Ddrive\Automation_Projects\SMA_ExcelGeneration\inputFile_5G3001_latest.xlsx"
#excelFile = "C:\Users\badrinat\Nokia\MN E2ETRS BTS Transport SCT - Script Modular Approach\IXIA_Conversion\inputFile_5G3001_latest.xlsx"
#rootDir = input('Path to features folder: ')
rootDir = "C:\Ddrive\Automation_Projects\SMA_ExcelGeneration\SRAN19"
#hw_type = input('Provide HW details ex:FSMR3/ASIK: ')
#rel_type = input('Provide release name ex:5G/SRAN/CBTS:')
hw_type = "ASIB"
rel_type = "SRAN"
main(rootDir)