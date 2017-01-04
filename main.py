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
    finded = ty in simple_type_def
    if finded:
        return True

    # 判断是否函数类型
    if ty.startswith(type_function):
        return True


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


# 读取xls文件内容，并过滤注释行
def read_cells_from_xls(filepath):
    workbook = xlrd.open_workbook(filepath)
    sheets = workbook.sheets()
    cells = []
    for sheet in sheets:
        if sheet.ncols <= 0:
            continue
        for y in range(0, sheet.nrows):
            text = sheet.cell_value(y, 0)
            if isinstance(text, unicode) and text.startswith('//'):
                continue
            cells.append(sheet.row(y))
    return cells


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

    # 添加元素项，数组类型name为None
    def add_member(self, name, pt):
        self.members.append((name, pt))

    def get_member(self, name):
        for key,value in self.members:
            if key == name:
                return value
        return None

    def colstr(self):
        if self.begin_col == self.end_col:
            return '(col %s)' % to_xls_col(self.begin_col)
        else:
            return '(col %s,%s)' % (to_xls_col(self.begin_col), to_xls_col(self.end_col))

    # 打印输出
    def tostr(self, identity):
        result = ''
        if self.type == type_array:
            result += self.type
            result += self.colstr()
            result += '[\n'
            for i in range(len(self.members)):
                (_, pt) = self.members[i]
                result += '\t' * (identity+1)
                result += pt.tostr(identity+1)
                result += ','
                result += '\n'
            result += '\t' * identity
            result += ']'
        elif self.type == type_dict:
            result += self.type
            result += self.colstr()
            result += '{\n'
            for i in range(len(self.members)):
                (name, pt) = self.members[i]
                result += '\t' * (identity+1)
                result += name
                result += pt.colstr()
                result += ': '
                result += pt.tostr(identity+1)
                result += ','
                result += '\n'
            result += '\t' * identity
            result += '}'
        else:
            result += self.type
            result += self.colstr()
        return result

    def __str__(self):
        return self.tostr(0)

    def parse(self, cell_row):
        if self.type == type_int:
            text = cell_row[self.begin_col]
            return str(int(text))
        elif self.type == type_string:
            text = cell_row[self.begin_col]
            return r"'%s'" % text
        else:
            return ''


# 解析器
class RootParser:
    def __init__(self, ts):
        self.tokens = ts
        self.pos = 0

        # 配置符合类型对应的解析函数
        self.parser_cfg = {
            type_array: self.build_array_node,
            type_dict: self.build_dict_node,
        }

        root_node = NodeParser(type_dict)
        root_node.begin_col = self.cur_token().col
        self.build_dict_node(root_node)
        root_node.end_col = self.tokens[self.pos-1].col
        self.root_node = root_node

    # 取得当前读取位置的符号
    def cur_token(self):
        if self.pos >= len(self.tokens):
            return None
        return self.tokens[self.pos]

    # 读取头前进
    def skip_token(self):
        self.pos += 1

    # 输出解析错误信息并退出程序
    def build_node_error(self, msg):
        print '[build node error]%s' % msg
        exit(1)

    # 解析简单类型
    def build_simple_node(self):
        cur = self.cur_token()
        pt = NodeParser(cur.ty)
        pt.begin_col = cur.col
        pt.end_col = cur.col
        return pt

    # 解析复合类型
    def build_complex_node(self, pt):
        cur = self.cur_token()
        for k,v in complex_token_def.items():
            if v[0] == cur.ty:
                value = NodeParser(k)
                value.begin_col = cur.col
                self.skip_token()
                self.parser_cfg[k](value)
                return value
        self.build_node_error('解析列%d的%s类型元素时,遇到了位于列%d的意外的符号\'%s\'' % (pt.begin_col, pt.type, cur.col, str(cur)))
        return None

    # 解析数组类型
    def build_array_node(self, arr_node):
        cur = self.cur_token()
        (_, end_token) = complex_token_def[type_array]
        while cur is not None:
            if cur.ty == end_token:
                arr_node.end_col = cur.col
                break

            if cur.ty == '':
                self.build_node_error('列%d项%s没有声明类型' % (cur.col, str(cur.id)))
                break
            elif is_simple_type(cur.ty):
                value = self.build_simple_node()
                self.skip_token()
            else:
                value = self.build_complex_node(arr_node)
                self.skip_token()
            if value is not None:
                arr_node.add_member(cur.id, value)
            cur = self.cur_token()

    # 解析字典类型
    def build_dict_node(self, dict_node):
        cur = self.cur_token()
        (_, end_token) = complex_token_def[type_dict]
        while cur is not None:
            if cur.ty == end_token:
                dict_node.end_col = cur.col
                break

            if cur.id == '':
                self.build_node_error('解析列%d的%s类型元素时,遇到位于列%d的元素缺少变量名' % (dict_node.begin_col, dict_node.type, cur.col))
                break

            if cur.ty == '':
                self.build_node_error('列%d项%s没有声明类型' % (cur.col, str(cur.id)))
                break
            elif is_simple_type(cur.ty):
                value = self.build_simple_node()
                self.skip_token()
            else:
                value = self.build_complex_node(dict_node)
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
                            self.build_node_error('path_id(%s) is invalid' % cur.id)
                            break
                    last_node.add_member(path_id[path_len-1], value)
                else:
                    dict_node.add_member(cur.id, value)
            cur = self.cur_token()


def build_node_parser(cells):
    ts = []
    for col in range(0, len(cells[0])):
        ty = cells[0][col].value
        iden = cells[1][col].value
        t = Token(ty, iden, col)
        ts.append(t)
    ps = RootParser(ts)
    return ps.root_node

if __name__ == '__main__':
    cells = read_cells_from_xls('test.xlsx')    # 过滤注释行
    root_parser = build_node_parser(cells)
    print root_parser




