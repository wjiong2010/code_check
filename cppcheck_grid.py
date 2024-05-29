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


def parse_work(m, node):
    for subNode in node.childNodes:
        if CR_ELEMENT_NODE == subNode.nodeType:
            if 'work_app' == subNode.nodeName:
                m.work_apps = subNode.getAttribute('item').split(',')
            elif 'work_module' == subNode.nodeName:
                m.work_modules = subNode.getAttribute('item').split(',')
            elif 'work_gl_app' == subNode.nodeName:
                m.work_gl_apps = subNode.getAttribute('item').split(',')
            elif 'work_pro_app' == subNode.nodeName:
                m.work_pro_apps = subNode.getAttribute('item').split(',')
            else:
                print("UNKNOWN node type: " + subNode.nodeName)


class Team:
    class Member:
        def __init__(self):
            self.name_en = ''
            self.name_cn = ''
            self.work_modules = []
            self.work_apps = []
            self.work_gl_apps = []
            self.work_pro_apps = []
            self.cr_result = []

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

        is_pro_track = False
        fw_track_tag = "framework\\track\\"
        len_fw_track_tag = len(fw_track_tag)
        print("tag_path: " + code_file_path)
        module_index = code_file_path.find(fw_track_tag)
        if module_index == -1:
            fw_track_tag = "framework\\trackPro\\"
            len_fw_track_tag = len(fw_track_tag)
            print("pro_tag_path: " + code_file_path)
            module_index = code_file_path.find(fw_track_tag)
            is_pro_track = True
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
            app_i = module_index + 5  # skip 'apps\'
            app_j = code_file_path[app_i:].find("\\")
            app_name = code_file_path[app_i:app_i+app_j].upper()
            print("app_i,j: " + str(app_i) + "," + str(app_j) + " app_name: " + app_name)
            if is_pro_track:
                for mb in self.members:
                    if app_name in mb.work_pro_apps:
                        mb.cr_result.append(cr_result)
            else:
                for mb in self.members:
                    if app_name in mb.work_apps:
                        mb.cr_result.append(cr_result)
        else:
            print("module_name: " + module_name)

        return 0

    def parse_developer(self, node):
        for subNode in node.childNodes:
            if CR_ELEMENT_NODE == subNode.nodeType:
                if 'developer' == subNode.nodeName:
                    m = self.Member()
                    m.name_en = subNode.getAttribute('name_en')
                    m.name_cn = subNode.getAttribute('name_cn')
                    parse_work(m, subNode)
                    self.members.append(m)

    def print_members(self):
        for mb in self.members:
            print(mb.name_cn)
            print(mb.name_en)
            print(mb.work_apps)
            print(mb.work_modules)
            print(mb.work_gl_apps)
            print(mb.work_pro_apps)

    def init_members(self):
        file = "team.xml"
        try:
            dom = xml.dom.minidom.parse(file)
        except DOMException:
            raise Exception(file + ' is NOT a well-formed XML file.')

        root = dom.documentElement

        if root.nodeName != 'team':
            raise Exception('XML has no \'Team\' element.')

        # Get attributes of results
        results_ver = root.getAttribute('version')
        print(results_ver)

        for node in root.childNodes:
            if CR_ELEMENT_NODE == node.nodeType:  # 1 is Element
                if 'application' == node.nodeName or 'system' == node.nodeName:
                    self.parse_developer(node)

    def __init__(self):
        self.members = []


software_develop_team = Team()


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
    software_develop_team.init_members()
    # software_develop_team.print_members()
    xmlparse_f("Code_Review.xml", software_develop_team)
    xmlparse_f("Code_Review_pro.xml", software_develop_team)
    software_develop_team.save_as_text("Code_Review.txt")


main()
