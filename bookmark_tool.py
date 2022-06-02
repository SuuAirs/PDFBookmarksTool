import sys
import os
import argparse
import re
#import code
import json

from pikepdf import Pdf, OutlineItem
from pikepdf import Array, Name, Page, String

if sys.version_info < ( 3, 7 ):
    raise NotImplementedError("pikepdf requires Python 3.7+")

# https://github.com/aliaafee/pypdfbookmarks
# https://github.com/RussellLuo/pdfbookmarker
# https://github.com/dnxbjyj/py-project/tree/master/AddPDFBookmarks
# https://github.com/Cluas/bookmark2pdf
# https://www.zhihu.com/question/344805337/answer/1116258929


class PublicFunc():
    @staticmethod
    def write_text_file(content, output_path, encoding='utf-8'):
        with open(output_path, 'w', encoding=encoding) as f:
            f.write(content)

    @staticmethod
    def read_text_file(input_path, encoding='utf-8'):
        with open(input_path, 'r', encoding=encoding) as f:
            return f.readlines()
        
    @staticmethod
    def read_json_file(path, encoding='utf8'):
        """
        读取json文件
        """
        
        with open(path, 'r', encoding=encoding) as f:
            return json.load(f)

    @staticmethod
    def write_json_file(path, data, encoding='utf8'):
        """
        写入json数据
        """
        
        with open(path, 'w', encoding='utf8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    
class Constant():
    '''Constant'''

    # MODE
    ADD = 'add'
    REMOVE = 'remove'
    FORMAT = 'format'
    EXPORT = 'export'

    DICT_OUT_EXT = {
        ADD: '.pdf',
        REMOVE: '.pdf',
        EXPORT: '.txt',
        FORMAT: '.txt'
    }

    # 书签标题与页码间的分隔符, 可使用多个
    MARK_PAGE = '\t'

    # 代表书签标题级别的符号, 可使用多个
    MARK_LEVEL = '\t'

    __re_special_chars = r'+-*\()[]{}^|'

    @classmethod
    def mark_process(cls):
        '''若MARK_PAGE、MARK_LEVEL为正则表达式的特殊字符，'''

        def _escape_mark(mark_str):
            char_list = list(mark_str)
            for i, char in enumerate(char_list):
                if char in cls.__re_special_chars:
                    char_list[i] = '\\' + char
            return ''.join(char_list)

        cls.MARK_PAGE_RE = _escape_mark(cls.MARK_PAGE)
        cls.MARK_LEVEL_RE = _escape_mark(cls.MARK_LEVEL)


Constant.mark_process()

class BookmarkNode(object):
    def __init__(self, level=1, title=None, page_num=0):
        self.parent = None
        self.child = []
        self.title = title
        self.page_num = int(page_num) if page_num != None else None
        self.level = int(level)
        return

    def add_child(self, child):
        """Add a child node of type BookmarkNode"""
        self.child.append(child)
        child.parent = self

    def set_parent(self, new_parent):
        """Set the parent of this node to a new one"""
        old_parent = self.parent
        new_parent.add_child(self)
        if old_parent is not None:
            old_parent.child.remove(self)

    def move_to(self, new_index):
        """Move this node to a new location in the list of children"""
        if self.parent is None:
            print("Cannot move the root node")
            return

        self.parent.child.insert(
            new_index,
            self.parent.child.pop(self.parent.child.index(self)))

    def remove(self):
        """Remove/Delete the Node"""
        if self.parent is None:
            print("Cannot remove the root node")

        self.parent.child.remove(self)

    def load_from_pdf(self, pdfreader):
        """Load bookmarks from PyPDF2 PdfFileReader"""
        
        def find_dest(ref, names):
            resolved = None
            if isinstance(ref, Array):
                resolved = ref[0]
            else:
                if names is None:
                    resolved = None
                else:
                    for n in range(0, len(names) - 1, 2):
                        if names[n] == ref:
                            if names[n+1]._type_name == 'array':
                                named_page = names[n+1][0]
                            elif names[n+1]._type_name == 'dictionary':
                                named_page = names[n+1].D[0]
                            else:
                                raise TypeError("Unknown type: %s" % type(names[n+1]))
                            resolved = named_page
                            break
            if resolved is not None:
                return Page(resolved).index
            else:
                return None
                
        def _getDestinationPageNumber(outline, names):
            if outline.destination is not None:
                if isinstance(outline.destination, Array):
                    # 12.3.2.2 Explicit destination
                    # [raw_page, /PageLocation.SomeThing, integer parameters for viewport]
                    raw_page = outline.destination[0]
                    try:
                        page = Page(raw_page)
                        dest = page.index
                    except:
                        dest = find_dest(outline.destination, names)
                elif isinstance(outline.destination, String):
                    # 12.3.2.2 Named destination, byte string reference to Names
                    # dest = f'<Named Destination in document .Root.Names dictionary: {outline.destination}>'
                    assert names is not None
                    dest = find_dest(outline.destination, names)
                elif isinstance(outline.destination, Name):
                    # 12.3.2.2 Named desintation, name object (PDF 1.1)
                    # dest = f'<Named Destination in document .Root.Dests dictionary: {outline.destination}>'
                    dest = find_dest(outline.destination, names)
                elif isinstance(outline.destination, int):
                    # Page number
                    dest = outline.destination
                else:
                    dest = outline.destination
                return dest
            else:
                try:
                    return find_dest(outline.action.D, names)
                except AttributeError:
                    return None

        def get_names(pdf):
            # https://github.com/pikepdf/pikepdf/issues/149#issuecomment-860073511
            def has_nested_key(obj, keys):
                ok = True
                to_check = obj
                for key in keys:
                    if key in to_check.keys():
                        to_check = to_check[key]
                    else:
                        ok = False
                        break
                return ok   
        
            if has_nested_key(pdf.Root, ['/Names', '/Dests']):
                obj = pdf.Root.Names.Dests
                names = []
                ks = obj.keys()
                if '/Names' in ks:
                    names.extend(obj.Names)
                elif '/Kids' in ks:
                    for k in obj.Kids:
                        names.extend(get_names(k))
                else:
                    assert False
                return names
            else:
                return None
            
        def _generate_tree(parent_node, cur_outline, level):
            
            current_node = BookmarkNode()
            current_node.title = cur_outline.title.strip()
            #current_node.page_num = int(cur_outline.destination._type_code)
            page_num = _getDestinationPageNumber(cur_outline, names)
            page_num = int(page_num + 1) if page_num != None else None # 如果有页码,则 + 1
            current_node.page_num = page_num
            current_node.level = int(level)
            parent_node.add_child(current_node)

            for child_outline in cur_outline.children:
                _generate_tree(current_node, child_outline, level+1)
        
        try: 
            names = get_names(pdfreader)
        except AttributeError:
            names = None
        
        with pdfreader.open_outline() as outline_obj:
            for root_outline in outline_obj.root:
                _generate_tree(self, root_outline, level=1)

    def add_to_pdf(self, pdf_obj):
        
        def _add_bookmark(cur_node, parent_outline):
            page_num = cur_node.page_num - 1 if cur_node.page_num != None else None
            cur_outline = OutlineItem(cur_node.title, page_num)
            if cur_node.level == 1:
                parent_outline.append(cur_outline)
            else:
                parent_outline.children.append(cur_outline)
                
            for child_node in cur_node.child:
                _add_bookmark(child_node, cur_outline)
                
        with pdf_obj.open_outline() as outline_obj:
            for child_node in self.child:
                _add_bookmark(child_node, outline_obj.root)    

    def load_from_txt(self, txt_file_path, encoding='utf-8'):
        bmk_text_lines = PublicFunc.read_text_file(txt_file_path, encoding=encoding)
        self.load_from_text(bmk_text_lines)
        
    def load_from_text(self, bmk_text_lines):

        def _make_up_parent_root(cur_level, cur_title, node_dict):
            prev_level = cur_level - 1
            if prev_level not in node_dict.keys():
                print(f'Warning: Title "{cur_title}": missing {prev_level:.0f} level title')
                _make_up_parent_root(prev_level, cur_title, node_dict)
                node_dict[prev_level] = BookmarkNode(title='.'*5, level=prev_level)
                node_dict[prev_level-1].add_child(node_dict[prev_level])
            else:
                return

        # 如果直接输入的是文字,则按行转为list
        if type(bmk_text_lines) is str:
            bmk_text_lines = bmk_text_lines.split('\n')
            
        offset = 0
        node_dict = {0: self}

        for line in bmk_text_lines:
            #line = line.strip(' ')
            # / / 后面填上 页码中的第一页对应PDF的第几个页面
            if line.strip().startswith('//'):
                try:
                    offset = int(line[2:].strip()) - 1
                except ValueError:
                    pass
                continue
            res = re.match(rf'^(({Constant.MARK_LEVEL_RE})*)(.*?)({Constant.MARK_PAGE_RE})(\d*)', line)
            if res:
                level_mark, _, title, _, page_num = res.groups()
                cur_level = len(level_mark) / len(Constant.MARK_LEVEL) + 1  # \t count stands for level
                if cur_level % 1: # if title level is not int
                    raise ValueError('Bookmark file not be formated!')
                page_num = int(page_num) + offset if page_num != '' else None
                cur_node = BookmarkNode(level=cur_level, title=title, page_num=page_num)

                _make_up_parent_root(cur_level, title, node_dict)

                node_dict[cur_level - 1].add_child(cur_node)
                node_dict[cur_level] = cur_node
                for i_ in list(node_dict.keys()):
                    if i_ > cur_level:
                        node_dict.pop(i_)

    def convert_to_txt(self):
        """Recursively print all the nodes of this tree"""

        def _outline_format(bookmark_list, node=None):
            if node is None:
                node = self
            else:
                level_mark = Constant.MARK_LEVEL * (node.level - 1)
                page_num = node.page_num if node.page_num != None else ''
                bookmark_txt = f'{level_mark}{node.title}{Constant.MARK_PAGE}{page_num}'
                bookmark_list.append(bookmark_txt)

            for num, child in enumerate(node.child):
                _outline_format(bookmark_list, node=child)

        bookmark_list = []
        _outline_format(bookmark_list)

        return '\n'.join(bookmark_list)

    def load_from_dict(self, bookmarks_dict, level=0):
        self.title = bookmarks_dict['title']
        page_num = bookmarks_dict['page_num']
        self.page_num =  page_num - 1 if page_num != '' else None
        self.child = []
        self.level = level
        #self.level = bookmarks_dict['level']
        for child_dict in bookmarks_dict['child']:
            child = BookmarkNode()
            self.add_child(child)
            child.load_from_dict(child_dict, level+1)

    def convert_to_dict(self):
        return {
            'title': self.title,
            'page_num': self.page_num + 1 if self.page_num != None else '',
            'child': [child.convert_to_dict() for child in self.child],
            'level': self.level
        }

    def load_from_json(self, json_file_path, encoding='utf-8'):
        bookmarks_dict = PublicFunc.read_json_file(json_file_path, encoding=encoding)
        self.load_from_dict(bookmarks_dict)

    def convert_to_json(self):
        return json.dumps(self.convert_to_dict(), ensure_ascii=False, indent=4)

    def print_tree(self, num=0, node=None, depth=0):
        """Recursively print all the nodes of this tree"""
        if node is None:
            node = self
        else:
            print("{}[{}] {}".format("   " * depth, num, node))
        for num, child in enumerate(node.child):
            self.print_tree(num, child, depth + 1)
    
    def print_tree2(self):
        print(self.convert_to_txt())
        
    def print_child(self):
        """Print all the children of this node"""
        for num, child in enumerate(self.child):
            print("[{}] {}".format(num, child))

    def __repr__(self):
        """String representation of object"""
        return "{} -> p{}{}".format(
            self.title,
            self.page_num,
            ", c{}".format(len(self.child)) if self.child else "")


class MyPDFHandler(object):
    '''
    封装的PDF文件处理类
    '''

    def __init__(self, in_pdf_path):
        self.__pdf_reader = Pdf.open(in_pdf_path, allow_overwriting_input=True)

    def generate_bookmark_tree(self, input=''):
        self.bookmark_tree = BookmarkNode(title='Root')

        if input == '':
            self.bookmark_tree.load_from_pdf(self.__pdf_reader)
            return

        if type(input) is dict:
            self.bookmark_tree.load_from_dict(input)
            return

        name_parts = os.path.splitext(input.lower())
        if name_parts[1] == '.txt':
            self.bookmark_tree.load_from_txt(input)
        elif name_parts[1] == '.json':
            self.bookmark_tree.load_from_json(input)
        else:
            raise Exception(f'Invalid input file: {input}')

    def bookmark_tree_to_text_file(self, out_bookmark_path, encoding='utf-8'):
        name_parts = os.path.splitext(out_bookmark_path)

        if name_parts[1].lower() == '.json':
            bookmark_txt = self.bookmark_tree.convert_to_json()
        else:
            bookmark_txt = self.bookmark_tree.convert_to_txt()

        PublicFunc.write_text_file(bookmark_txt, out_bookmark_path, encoding)

    def remove_bookmarks(self):
        with self.__pdf_reader.open_outline() as outline_obj:
            outline_obj.root.clear()
        
    def add_bookmarks_to_pdf(self):
        self.bookmark_tree.add_to_pdf(self.__pdf_reader)

    def write_to_pdf(self, out_pdf_path):
        self.__pdf_reader.save(out_pdf_path)

    @staticmethod
    def format_bookmark_file(input_bmk_path,
                             output_bmk_path,
                             in_encoding='utf-8',
                             out_encoding='utf-8'):

        # 读取书签文件, 每行为列表的一个元素
        bmk_text_lines = PublicFunc.read_text_file(input_bmk_path, in_encoding)

        list_reg_patern = [
            
            # 不以数字开头的行，例如：前言
            (r'^(%s|\s)*([^\d%s第])' % (Constant.MARK_LEVEL_RE, Constant.MARK_LEVEL_RE),
             Constant.MARK_LEVEL*0 + r'\2'),
            
            # 一级标题：第x章
            (r'^(%s|\s)*(第\d{1,}章)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),

            # 一级标题：第x章
            (r'^(%s|\s)*(第[一二三四五六七八九十〇IV]+章)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),
            
            # 一级标题 Chapter 1
            (r'^(%s|\s)*(chapter\s*[\dIV]+)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),
            

            # 一级标题：1标题  或  1. 标题
            (r'^(%s|\s)*(\d{1,}\.?)\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),
            
            # 二级标题：1.1标题
            (r'^(%s|\s)*(\d{1,}\.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL + r'\2 '),

            # 二级标题：第x节
            (r'^(%s|\s)*(第[一二三四五六七八九十IV]+节)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL + r'\2 '),

            # 三级标题 1.1.1
            (r'^(%s|\s)*(\d{1,}\.\d{1,}\.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL*2 + r'\2 '),

            # 四级标题 1.1.1.1
            (r'^(%s|\s)*(\d{1,}\.\d{1,}\.\d{1,}.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL*3 + r'\2 '),

            # 标题与页码间:  18   或  18-25
            (r'(%s|\s)*(\d{1,})-?\d{0,}\s*\r?$' % Constant.MARK_PAGE_RE,
             Constant.MARK_PAGE + r'\2'),    # 页码
            
            # 没有页码的标题, 末尾加上页码分隔符
            (r'([^\d])(%s|\s)*$' % Constant.MARK_PAGE,
             r'\1' + Constant.MARK_PAGE),    # 页码
        ]

        res_txt = ''.join(bmk_text_lines)
        for reg_txt, rep_txt in list_reg_patern:
            res_txt = re.sub(reg_txt, rep_txt, res_txt, flags=re.M)

        PublicFunc.write_text_file(res_txt, output_bmk_path, out_encoding)


def get_cmd_args():

    dest_str = ('pdf bookmark tool.\n'
                'Attention: Paths containing spaces must be enclosed in double quotes')
    parser = argparse.ArgumentParser(description=dest_str)
    parser.add_argument('-mode', dest='mode', default='add',
                        choices=list(Constant.DICT_OUT_EXT.keys()),
                        action='store',
                        help='add, remove, export, format.')
    parser.add_argument('-i', dest='i', action='store',
                        help='origin pdf filename.')
    parser.add_argument('-bmk', dest='bmk', action='store',
                        help='bookmarks file.')
    parser.add_argument('-o', dest='o', action='store',
                        help='save to filename.')
    parser.add_argument('-y', dest='overwrite', action='store_true',
                        help='overwrite output file if it already exists')

    args = parser.parse_args()

    # for test
    if not args.i:
        # args.i = os.getcwd() + '/examples/book-new.pdf'
        # args.mode = 'export'
        # args.i = os.getcwd() + '/exa  mples/book-new_remove.pdf'
        
        args.i = r'E:/02-书籍/2Python/Python核心编程(第3版)中文版.pdf'
        #args.bmk = r'E:\test_pdf - 副本\NPC三电平大功率PWM整流器预测控制研究_张旭.json'
        #args.o = 'E:/浏览器下载/基于MPC与预瞄理论的自动驾驶车辆轨迹跟随控制研究_马瀚森1.pdf'
        # args.i = 'E:/浏览器下载/基于模型预测控制的无人车辆轨迹跟踪研究_杨正龙.pdf'
        args.mode = Constant.ADD
        args.overwrite = True

    veryfy_args(args)
    
    return args


def veryfy_args(args):

    # if args.mode not in Constant.DICT_OUT_EXT.keys():
    #     print(f'ERROR: Unknow mode: {args.mode}')
    #     exit(2)
    args.mode = args.mode.lower()
    
    # veryfy input path
    if not args.i:
        print('ERROR: Input file not be specified!')
        sys.exit(2)

    if not os.path.exists(args.i):
        print(f'ERROR: Input file not exist: {args.i}')
        sys.exit(2)

    input_path_parts = os.path.splitext(args.i)

    if ((args.mode in [Constant.ADD, Constant.REMOVE, Constant.EXPORT])
            and input_path_parts[1].lower() != '.pdf'):
        print(f'In mode "{args.mode}", Input file must be PDF format!')
        sys.exit(2)

    # veryfy output path
    if not args.o:
        args.o = f'{input_path_parts[0]}{Constant.DICT_OUT_EXT[args.mode]}'
        # args.o = f'{input_path_parts[0]}_{args.mode}{Constant.DICT_OUT_EXT[args.mode]}'

    output_path_parts = os.path.splitext(args.o.lower())
    if (args.mode in [Constant.ADD, Constant.REMOVE]) and output_path_parts[1] != '.pdf':
        args.o = args.o + '.pdf'

    if (args.mode in [Constant.EXPORT, Constant.FORMAT]) and output_path_parts[1] == '':
        args.o = args.o + '.txt'

    if (not args.overwrite) and (os.path.exists(args.o)):
        user_choice = input(f'Destnation file: {args.o} \n already exists, overwrite? (y/n)')
        if user_choice.lower() != 'y':
            sys.exit(1)

    # veryfy bookmark path
    if args.mode == Constant.ADD:
        if (not args.bmk):
            #args.bmk = input('input_bookmark_file_path: ')
            args.bmk = input_path_parts[0] + '.txt'

        if not os.path.exists(args.bmk):
            print(f'ERROR: Bookmark file not exist: {args.bmk} ')
            sys.exit(2)


if __name__ == "__main__":

    args = get_cmd_args()
    
    print('\n' + '-'*30)
    print(f'Process mode: \t{args.mode}')
    print(f'Input file: \t{args.i}')
    if args.mode == Constant.ADD:
        print(f'Bookmark file: \t{args.bmk}')
    print(f'Output file: \t{args.o}')

    if args.mode == Constant.ADD:
        pdf_handler = MyPDFHandler(args.i)
        print("Read origin pdf file success...")

        pdf_handler.generate_bookmark_tree(args.bmk)
        pdf_handler.remove_bookmarks()
        #pdf_handler.bookmark_tree.print_tree2()
        pdf_handler.add_bookmarks_to_pdf()
        print("Parse bookmark success...")

        pdf_handler.write_to_pdf(args.o)
        print("Save pdf with bookmark success...")
        
    elif args.mode == Constant.REMOVE:
        pdf_handler = MyPDFHandler(args.i)
        pdf_handler.remove_bookmarks()
        pdf_handler.write_to_pdf(args.o)
        print("Remove bookmarks success...")

    elif args.mode == Constant.EXPORT:
        pdf_handler = MyPDFHandler(args.i)
        pdf_handler.generate_bookmark_tree()
        pdf_handler.bookmark_tree_to_text_file(args.o)
        print("Export bookmarks success...")
        
    elif args.mode == Constant.FORMAT:
        MyPDFHandler.format_bookmark_file(args.i, args.o)
        print("Format bookmarks success...")
    
    print('-'*30 + '\n')
    
    
# def shell():
#     code.interact(
#         banner="PyPDF Bookmarks Shell, pypdfbm_help() for help",
#         local=globals()
#     )
