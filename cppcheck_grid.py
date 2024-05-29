import xml.dom.minidom
from xml.dom import DOMException

import xlwt
import xlrd

CR_ELEMENT_NODE = 1


class ResultLocation:
    def __init__(self):
        self.filename = ''
        self.line = ''
        self.column = ''
        self.info = ''


class CodeReviewResult:
    def __init__(self):
        self.id = ''
        self.severity = ''
        self.msg = ''
        self.verbose = ''
        self.cwe = ''
        self.file0 = ''
        self.locations = []


class Member:
    def __init__(self):
        self.name_en = ''
        self.name_cn = ''
        self.work_modules = []
        self.work_apps = []
        self.cr_result = []


class Team:
    def save_as_text(self, txt):
        summary_in_text = ''

        if len(self.members) == 0:
            print('No members found.')
            return -1

        for mb in self.members:
            summary_in_text += f"{mb.name_cn}({mb.name_en}):\n"
            for cr in mb.cr_result:
                summary_in_text += f"\t{cr.severity.upper()}:\n"
                summary_in_text += f"\t{cr.msg}\n"
                if len(cr.msg) != len(cr.verbose):
                    summary_in_text += f"\t{cr.verbose}\n"
                for location in cr.locations:
                    if len(cr.locations) == 1:
                        summary_in_text += f"\t\t{cr.id}:\n"
                    else:
                        summary_in_text += f"\t\t{location.info}:\n"
                    summary_in_text += f"\t\t{location.filename} line:{location.line}, col:{location.column}\n"
                summary_in_text += f"\t修复结果：\n"
                summary_in_text += '\n'
            summary_in_text += '\n\n'

        with open(txt, 'w', encoding='utf-8') as f:
            f.write(summary_in_text)

    def arrange_results(self, cr_result):
        if len(cr_result.file0) != 0:
            code_file_path = cr_result.file0
            code_file_path = code_file_path.replace('/', '\\')
        elif len(cr_result.locations) != 0 and cr_result.locations[0].filename != '':
            code_file_path = cr_result.locations[0].filename
        else:
            code_file_path = ''

        fw_track_tag = "framework\\track\\"
        len_fw_track_tag = len(fw_track_tag)
        print("tag_path: " + code_file_path)
        module_index = code_file_path.find(fw_track_tag)
        if module_index == -1:
            fw_track_tag = "framework\\trackPro\\"
            len_fw_track_tag = len(fw_track_tag)
            print("pro_tag_path: " + code_file_path)
            module_index = code_file_path.find(fw_track_tag)
            if module_index == -1:
                return -1

        module_index += len_fw_track_tag
        # get module name from file0
        module_name = code_file_path[module_index:module_index + 6]
        module_name = module_name.upper()
        print("module_index: " + str(module_index) + "  module_name: " + module_name)
        if 'IO' == module_name[:2] or \
                'GPS' == module_name[:3] or \
                'MCU' == module_name[:3] or \
                'BLE' == module_name[:3] or \
                'DBG' == module_name[:3] or \
                'COMM' == module_name[:4] or \
                'ATCI' == module_name[:4] or \
                'PROT' == module_name[:4] or \
                'UTILS' == module_name[:5] or \
                'SERIAL' == module_name[:6] or \
                'SENSOR' == module_name[:6]:
            i = module_name.rfind('\\')
            if -1 != i:
                module_name = module_name[:i]
            print("proc module_name: " + module_name + " i:" + str(i))
            for mb in self.members:
                if module_name in mb.work_modules:
                    mb.cr_result.append(cr_result)
        elif 'APPSGL' == module_name[:6]:
            app_i = module_index + 7  # skip 'appsGL/'
            app_name = code_file_path[app_i:app_i + 3].upper()
            print("gl_app_i: " + str(app_i) + " gl_app_name: " + app_name)
            for mb in self.members:
                if app_name in mb.work_gl_apps:
                    mb.cr_result.append(cr_result)
        elif 'APPS' == module_name[:4]:
            app_i = module_index + 5  # skip 'apps/'
            app_name = code_file_path[app_i:app_i + 3].upper()
            print("app_i: " + str(app_i) + " app_name: " + app_name)
            for mb in self.members:
                if app_name in mb.work_apps:
                    mb.cr_result.append(cr_result)
        else:
            print("module_name: " + module_name)

        return 0

    def init_members(self):
        for wk in self.work_arrange:
            m = Member()
            m.name_en = wk['name_en']
            m.name_cn = wk['name_cn']
            m.work_apps = wk['work_app'].split(',')
            m.work_modules = wk['work_module'].split(',')
            m.work_gl_apps = wk['work_gl_app'].split(',')
            self.members.append(m)

    def __init__(self):
        self.members = []
        self.work_arrange = [
            {'name_en': "Len Liu", "name_cn": "刘信", 'work_app': 'AIS,EPS,EFS,FSC,UFS,JDC,JBS', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Claire Liu", "name_cn": "刘慧", 'work_app': 'BID,RTO,PDS,SOS,SSR', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Aleo Liu", "name_cn": "刘洋洋", 'work_app': 'FFC,FRI,EMG,CRA,ASC,HBM,RAS,RCS,GNA',
             'work_module': 'GPS,PROT', 'work_gl_app': ''},
            {'name_en': "Harper Kuang", "name_cn": "匡婷", 'work_app': 'SIM,SPD,TOW,OWH,PIN,TMA,WLT',
             'work_module': '', 'work_gl_app': ''},
            {'name_en': "Rain Wu", "name_cn": "吴瑞", 'work_app': 'DTT,MON,RMD,DAT,TKS,UDT', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Vincent Cui", "name_cn": "崔子晨", 'work_app': 'IDA,CDA,GEO,PEO,PEG,GAM,FKS,NMD',
             'work_module': '', 'work_gl_app': 'CFG,NMD,DOG,TMA,SFM,FKS,EMS,RDF'},
            {'name_en': "Bennett Cui", "name_cn": "崔斌", 'work_app': 'DIS,IOB,OUT,DOS,GDO,ACD,IEX,OEX,TMP,HUM,SLM',
             'work_module': 'IO', 'work_gl_app': 'GEO,JDC,TEM,FRI,UPC,UPD,FVR'},
            {'name_en': "Haze Zhang", "name_cn": "张仲俊", 'work_app': 'SPA,BZA,DUC,CMD,UDF,AUS', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Allen Zhang", "name_cn": "张学忠", 'work_app': 'BSI,QSS,SRI,FVR,UPC,UPD,MDT',
             'work_module': '', 'work_gl_app': ''},
            {'name_en': "Abert Xu", "name_cn": "徐黎明", 'work_app': 'MQT,AVS,VVS,VMS,SMS,OWL,LTP,TLS',
             'work_module': '', 'work_gl_app': 'BSI,QSS,SRI,MQT,TLS,LTP,RTP'},
            {'name_en': "Bear Cao", "name_cn": "曹政", 'work_app': 'HMC,HRM,CDS,MSI,MSF', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Archie Li", "name_cn": "李叶齐", 'work_app': 'BTS,ROS,BAS,AEX,SVR', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Arthur Lee", "name_cn": "李永乐", 'work_app': '', 'work_module': '', 'work_gl_app': ''},
            {'name_en': "Jack Li", "name_cn": "李仁杰", 'work_app': '', 'work_module': '', 'work_gl_app': ''},
            {'name_en': "Elvin Shen", "name_cn": "沈子扬", 'work_app': 'CAN,CFU,TTR,CLT,AES,MQT',
             'work_module': 'COMM,SERIAL', 'work_gl_app': ''},
            {'name_en': "Ying Xiong", "name_cn": "熊鹰", 'work_app': 'CFG,CMS,TAP', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Ernie Hu", "name_cn": "胡心月", 'work_app': '', 'work_module': '', 'work_gl_app': ''},
            {'name_en': "Todd Zheng", "name_cn": "郑功良", 'work_app': 'DMS,DOG,FTP,FSC,FSI,FSS,PMS', 'work_module': '',
             'work_gl_app': ''},
            {'name_en': "Noah Qin", "name_cn": "秦伟", 'work_app': '', 'work_module': 'ATCI,SENSOR', 'work_gl_app': ''},
            {'name_en': "Swain Shen", "name_cn": "申亚", 'work_app': '', 'work_module': 'MCU', 'work_gl_app': ''},
            {'name_en': "Rmyh Hong", "name_cn": "洪飞", 'work_app': '', 'work_module': 'BLE', 'work_gl_app': ''},
            {'name_en': "Frank Sun", "name_cn": "孙旭", 'work_app': '', 'work_module': 'DBG', 'work_gl_app': ''},
            {'name_en': "Shannon Su", "name_cn": "苏竞成", 'work_app': '', 'work_module': 'UTILS', 'work_gl_app': ''}
        ]


app_team = Team()


# get attribute of node,
def __parse_attr(node, attr):
    value = node.getAttribute(attr)
    # not found will trigger exception, 'element xxx has no attributes.'
    if "" == value:
        print('Element \'' + node.nodeName + '\' has no \'' + attr + '\' attribute.')
    return value


def __parse_location(node, locs):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            loc = ResultLocation()
            if 'location' == subNode.nodeName:
                loc.filename = __parse_attr(subNode, 'file')
                loc.line = __parse_attr(subNode, 'line')
                loc.column = __parse_attr(subNode, 'column')
                loc.info = __parse_attr(subNode, 'info')
                locs.append(loc)


def __parse_errors(node, app_tm):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            cr_res = CodeReviewResult()
            if 'error' == subNode.nodeName:
                cr_res.id = __parse_attr(subNode, 'id')
                cr_res.severity = __parse_attr(subNode, 'severity')
                cr_res.msg = __parse_attr(subNode, 'msg')
                cr_res.verbose = __parse_attr(subNode, 'verbose')
                cr_res.file0 = __parse_attr(subNode, 'file0')
                cr_res.cwe = __parse_attr(subNode, 'cwe')
                __parse_location(subNode, cr_res.locations)
                print("len of locations: " + str(len(cr_res.locations)))
                app_tm.arrange_results(cr_res)


def xmlparse_f(file, app_tm):
    try:
        dom = xml.dom.minidom.parse(file)
    except DOMException:
        raise Exception(file + ' is NOT a well-formed XML file.')

    root = dom.documentElement

    if root.nodeName != 'results':
        raise Exception('XML has no \'Project\' element.')

    # Get attributes of results
    results_ver = __parse_attr(root, 'version')
    print(results_ver)

    # for mb in app_tm.members:
    #     print(mb.name_cn + '(' + mb.name_en + ')' + ' ')

    for node in root.childNodes:
        if CR_ELEMENT_NODE == node.nodeType:  # 1 is Element
            if 'errors' == node.nodeName:
                __parse_errors(node, app_tm)


def main():
    app_team.init_members()
    xmlparse_f("Code_Review.xml", app_team)
    app_team.save_as_text("Code_Review.txt")


main()
