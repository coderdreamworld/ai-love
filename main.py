# coding:utf-8
import inspect
import re
import xlrd

# 定义符号
type_array = 'array'
type_dict = 'dict'
type_int = 'int'
type_number = 'number'
type_string = 'string'
type_function = 'function'
type_table = 'table'

# 定义简单类型
simple_type_def = {
    type_int: True,
    type_number: True,
    type_string: True,
    type_table: True,
}

# 定义复合类型
complex_token_def = {
    type_array: ('array<', '>'),
    type_dict: ('dict<', '>'),
}


# 判断是否简单类型
def is_simple_type(ty):
    return ty in simple_type_def


# 将整数转换为excel的列
def to_xls_col(num):
    A = ord('A')
    result = ''
    tmp = num
    while True:
        factor = tmp / 26
        remain = tmp % 26
        tmp = factor
        result = chr(remain + A) + result
        if factor == 0:
            break
    return result


# 表头符号，每个符号由类型和名字组成，用于构造表格数据结构
class Token:
    def __init__(self, ty, iden, col):
        self.ty = ty
        self.id = iden
        self.col = col

    def __str__(self):
        if self.ty == '':
            return self.id
        else:
            return '%s:%s' % (self.id, self.ty)


# 表格数据结构，基于token流构造
class NodeParser:
    def __init__(self, type):
        self.type = type
        self.members = []
        self.begin_col = -1
        self.end_col = -1
        self.text = None

    # 添加元素项，数组类型name为None
    def add_member(self, name, pt):
        self.members.append((name, pt))

    def get_member(self, name):
        for key,value in self.members:
            if key == name:
                return value
        return None

    def eval(self, cell_row):
        if self.type == type_int:
            val = cell_row[self.begin_col].value
            if val == '':
                return 'nil'
            i = int(val)
            return str(i)
        elif self.type == type_string:
            text = cell_row[self.begin_col].value
            return '"' + text + '"'
        elif self.type == type_number:
            val = cell_row[self.begin_col].value
            return str(val)
        elif self.type == type_table:
            val = cell_row[self.begin_col].value
            return '{' + val + '}'
        elif self.type == type_function:
            val = cell_row[self.begin_col].value
            return '%s return %s end' % (self.text, val)
        elif self.type == type_array:
            result = '{'
            mcount = len(self.members)
            for i in range(mcount):
                (_,m) = self.members[i]
                result += m.eval(cell_row)
                if i <= mcount:
                    result += ','
            result += '}'
            return result
        elif self.type == type_dict:
            result = '{'
            mcount = len(self.members)
            for i in range(mcount):
                (id,m) = self.members[i]
                result += '%s=%s' % (id, m.eval(cell_row))
                if i <= mcount:
                    result += ','
            result += '}'
            return result
        else:
            return 'nil'


# 解析器
class SheetParser:
    def __init__(self, ts):
        self.tokens = ts
        self.pos = 0

        # 配置符合类型对应的解析函数
        self.parser_cfg = {
            type_array: self.parse_array_node,
            type_dict: self.parse_dict_node,
        }

        root_node = NodeParser(type_dict)
        root_node.begin_col = self.cur_token().col
        self.parse_dict_node(root_node)
        root_node.end_col = self.tokens[self.pos-1].col
        self.parser_tree_root = root_node

    # 取得当前读取位置的符号
    def cur_token(self):
        if self.pos >= len(self.tokens):
            return None
        return self.tokens[self.pos]

    # 读取头前进
    def skip_token(self):
        self.pos += 1

    # 输出解析错误信息并退出程序
    def parse_error(self, msg):
        print '[parser error]%s' % msg
        exit(1)

    # 解析简单类型
    def parse_simple_node(self):
        cur = self.cur_token()
        pt = NodeParser(cur.ty)
        pt.begin_col = cur.col
        pt.end_col = cur.col
        return pt

    # 解析复合类型
    def parse_complex_node(self, pt):
        cur = self.cur_token()
        for k,v in complex_token_def.items():
            if v[0] == cur.ty:
                value = NodeParser(k)
                value.begin_col = cur.col
                self.skip_token()
                self.parser_cfg[k](value)
                return value
        self.parse_error('解析列%d的%s类型元素时,遇到了位于列%d的意外的符号\'%s\'' % (pt.begin_col, pt.type, cur.col, str(cur)))
        return None

    # 解析数组类型
    def parse_array_node(self, arr_node):
        cur = self.cur_token()
        (_, end_token) = complex_token_def[type_array]
        while cur is not None:
            if cur.ty == end_token:
                arr_node.end_col = cur.col
                break

            if cur.ty == '':
                self.parse_error('列%d项%s没有声明类型' % (cur.col, str(cur.id)))
                break
            elif is_simple_type(cur.ty):
                value = self.parse_simple_node()
                self.skip_token()
            elif cur.ty.startswith(type_function):
                value = NodeParser(cur.ty)
                value.begin_col = cur.col
                value.end_col = cur.col
                value.text = cur.ty
                self.skip_token()
            else:
                value = self.parse_complex_node(arr_node)
                self.skip_token()
            if value is not None:
                arr_node.add_member(cur.id, value)
            cur = self.cur_token()

    # 解析字典类型
    def parse_dict_node(self, dict_node):
        cur = self.cur_token()
        (_, end_token) = complex_token_def[type_dict]
        while cur is not None:
            if cur.ty == end_token:
                dict_node.end_col = cur.col
                break

            if cur.id == '':
                self.parse_error('解析列%d的%s类型元素时,遇到位于列%d的元素缺少变量名' % (dict_node.begin_col, dict_node.type, cur.col))
                break

            if cur.ty == '':
                self.parse_error('列%d项%s没有声明类型' % (cur.col, str(cur.id)))
                break
            elif is_simple_type(cur.ty):
                value = self.parse_simple_node()
                self.skip_token()
            elif cur.ty.startswith(type_function):
                value = NodeParser(cur.ty)
                value.begin_col = cur.col
                value.end_col = cur.col
                value.text = cur.ty
                self.skip_token()
            else:
                value = self.parse_complex_node(dict_node)
                self.skip_token()
            if value is not None:
                # 路径变量需要解析取得操作对象
                path_id = cur.id.split('.')
                path_len = len(path_id)
                if path_len > 1:
                    last_node = dict_node
                    for i in range(path_len-1):
                        key = path_id[i]
                        last_node = last_node.get_member(key)
                        if last_node is None or last_node.type != type_dict:
                            self.parse_error('path_id(%s) is invalid' % cur.id)
                            break
                    last_node.add_member(path_id[path_len-1], value)
                else:
                    dict_node.add_member(cur.id, value)
            cur = self.cur_token()


# 读取xls文件内容，并过滤注释行
def read_sheets_from_xls(file_path):
    workbook = xlrd.open_workbook(file_path)
    sheets = []
    for sheet in workbook.sheets():
        if sheet.ncols <= 0:
            continue
        cells = []
        for y in range(0, sheet.nrows):
            # 过滤全空白行
            all_empty = True
            for v in sheet.row_values(y):
                if v != '':
                    all_empty = False
                    break
            if all_empty:
                continue
            text = sheet.cell_value(y, 0)
            # 过滤注释行
            if isinstance(text, unicode) and text.startswith('//'):
                continue
            cells.append(sheet.row(y))
        if len(cells) > 0:
            sheets.append((sheet.name, cells))
    return sheets


def build_parser_tree(sheet_cells):
    ts = []
    for col in range(len(sheet_cells[0])):
        ty = sheet_cells[0][col].value
        iden = sheet_cells[1][col].value
        t = Token(ty, iden, col)
        ts.append(t)
    sp = SheetParser(ts)
    return sp.parser_tree_root


if __name__ == '__main__':
    sheets = read_sheets_from_xls('test.xlsx')    # 过滤注释行
    for name, cells in sheets:
        parser_tree = build_parser_tree(cells)
        id_node = parser_tree.get_member('id')
        for y in range(2, len(cells)):
            row = cells[y]
            print '[%s]=%s,' % (id_node.eval(row), parser_tree.eval(row))




