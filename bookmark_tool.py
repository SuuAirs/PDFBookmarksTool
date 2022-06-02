import sys
import os
import argparse
import re
#import getopt
#import code
import json

from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from PyPDF2.generic import ByteStringObject, TextStringObject
from PyPDF2.generic import NameObject, createStringObject
from PyPDF2.errors import PdfReadError

# https://github.com/aliaafee/pypdfbookmarks
# https://github.com/RussellLuo/pdfbookmarker
# https://github.com/dnxbjyj/py-project/tree/master/AddPDFBookmarks
# https://github.com/Cluas/bookmark2pdf


class PublicFunc():
    @staticmethod
    def write_text_file(content, output_path, encoding='utf-8'):
        with open(output_path, 'w', encoding=encoding) as f:
            f.write(content)

    @staticmethod
    def read_text_file(input_path, encoding='utf-8'):
        with open(input_path, 'r', encoding=encoding) as f:
            return f.readlines(), f.read()


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

    # Use this if bookmark encoding could not be guessed
    FALLBACK_ENCODING = "utf-8"

    # 保留源PDF文件的所有内容和信息，在此基础上修改
    PDF_COPY = 'copy'
    # 仅保留源PDF文件的页面内容，在此基础上修改
    PDF_NEWLY = 'newly'

    # 书签标题与页码间的分隔符
    MARK_PAGE = '\t'

    # 代表书签标题级别的符号
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
        self.page_num = page_num
        self.level = level
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
        def _setup_page_id_to_num(pdf, pages=None, _result=None, _num_pages=None):
            if _result is None:
                _result = {}
            if pages is None:
                _num_pages = []
                pages = pdf.trailer["/Root"].getObject()["/Pages"].getObject()
            t = pages["/Type"]
            if t == "/Pages":
                for page in pages["/Kids"]:
                    _result[page.idnum] = len(_num_pages)
                    _setup_page_id_to_num(pdf, page.getObject(), _result, _num_pages)
            elif t == "/Page":
                _num_pages.append(1)
            return _result

        def _generate_tree(node, outlines, level):
            current_node = None
            for item in outlines:
                if type(item) is not list:
                    current_node = BookmarkNode()
                    if type(item.title) is ByteStringObject:
                        current_node.title = item.title.decode(Constant.FALLBACK_ENCODING, "backslashreplace")
                    else:
                        current_node.title = str(item.title)
                    current_node.title = current_node.title.strip()
                    current_node.page_num = pg_id_num_map[item.page.idnum] + 1
                    current_node.level = level
                    node.add_child(current_node)
                else:
                    _generate_tree(current_node, item, level=level+1)

        pg_id_num_map = _setup_page_id_to_num(pdfreader)
        _generate_tree(self, pdfreader.getOutlines(), level=1)

    def add_to_pdf(self, pdfwriter):
        """Save this bookmarks tree structure to PyPDF2 PdfFileWriter"""
        def _add_bookmark(node, pdfwriter, parent=None):
            pdf_node = None
            if node.parent is not None:
                pdf_node = pdfwriter.addBookmark(node.title, node.page_num - 1, parent=parent)
            for child in node.child:
                _add_bookmark(child, pdfwriter, pdf_node)
        _add_bookmark(self, pdfwriter)

    def load_from_txt(self, txt_file_path, encoding='utf-8'):

        def _make_up_parent_root(cur_level, node_dict):
            prev_level = cur_level - 1
            if prev_level not in node_dict.keys():
                _make_up_parent_root(prev_level-1, node_dict)
                node_dict[prev_level] = BookmarkNode(title=' ', level=prev_level)
                node_dict[prev_level-1].add_child(node_dict[prev_level])
            else:
                return

        offset = 0
        node_dict = {0: self}

        bmk_txt_lines, _ = PublicFunc.read_text_file(txt_file_path, encoding=encoding)

        for line in bmk_txt_lines:
            line = line.strip(' ')
            # / / 后面填上 页码中的第一页对应PDF的第几个页面
            if line.startswith('//'):
                try:
                    offset = int(line[2:].strip()) - 1
                except ValueError:
                    pass
                continue
            res = re.match(rf'^({Constant.MARK_LEVEL_RE}*)(.*?){Constant.MARK_PAGE_RE}(\d+)', line)
            if res:
                level_mark, title, page_num = res.groups()
                cur_level = len(level_mark) + 1  # \t count stands for level
                page_num = int(page_num) + offset
                cur_node = BookmarkNode(level=cur_level, title=title, page_num=page_num)

                _make_up_parent_root(cur_level, node_dict)

                node_dict[cur_level - 1].add_child(cur_node)
                node_dict[cur_level] = cur_node

    def convert_to_txt(self):
        """Recursively print all the nodes of this tree"""

        def _outline_format(bookmark_list, node=None):
            if node is None:
                node = self
            else:
                level_mark = Constant.MARK_LEVEL * (node.level - 1)
                bookmark_txt = f'{level_mark}{node.title}{Constant.MARK_PAGE}{node.page_num}'
                bookmark_list.append(bookmark_txt)

            for num, child in enumerate(node.child):
                _outline_format(bookmark_list, node=child)

        bookmark_list = []
        _outline_format(bookmark_list)

        return '\n'.join(bookmark_list)

    def load_from_dict(self, bookmarks_dict):
        self.title = bookmarks_dict['title']
        self.page_num = bookmarks_dict['page_num'] - 1
        self.child = []
        for child_dict in bookmarks_dict['child']:
            child = BookmarkNode()
            self.add_child(child)
            child.load_from_dict(child_dict)

    def convert_to_dict(self):
        return {
            'title': self.title,
            'page_num': self.page_num + 1,
            'child': [child.convert_to_dict() for child in self.child]
        }

    def load_from_json(self, json_file_path, encoding='utf-8'):
        _, bookmarks_json = PublicFunc.read_text_file(json_file_path, encoding=encoding)

        bookmarks_dict = json.loads(bookmarks_json)
        self.load_from_dict(bookmarks_dict)

    def convert_to_json(self):
        return json.dumps(self.convert_to_dict(), indent=4)

    def print_tree(self, num=0, node=None, depth=0):
        """Recursively print all the nodes of this tree"""
        if node is None:
            node = self
        else:
            print("{}[{}] {}".format("   " * depth, num, node))
        for num, child in enumerate(node.child):
            self.print_tree(num, child, depth + 1)

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

    def __init__(self, in_pdf_path, mode=Constant.PDF_COPY):
        '''
        用一个PDF文件初始化
        :param in_pdf_path: PDF文件路径
        :param mode: 处理PDF文件的模式，默认为Constant.PDF_COPY模式
        '''
        self.read_pdf(in_pdf_path, mode)

    def read_pdf(self, in_pdf_path, mode=Constant.PDF_COPY):
        
        def reset_eof_of_pdf_return_stream(pdf_stream_in:list):
            # find the line position of the EOF
            for i, x in enumerate(pdf_stream_in[::-1]):
                if b'%%EOF' in x:
                    actual_line = len(pdf_stream_in)-i
                    print(f'EOF found at line position {-i} = actual {actual_line}, with value {x}')
                    break

            # return the list up to that point
            return pdf_stream_in[:actual_line]

        # opens the file for reading
        #with open(in_pdf_path, 'rb') as p:
            #txt = (p.readlines())
            
        # get the new list terminating correctly
        #txtx = reset_eof_of_pdf_return_stream(txt)
        
        # write to new pdf
        #with open(in_pdf_path, 'wb') as f:
            #f.writelines(txtx)
        # 只读的PDF对象
        self.mode = mode

        self.__pdf_reader = PdfFileReader(in_pdf_path, strict=False)
        #self.__pdf_reader = PdfFileReader(p, strict=False)

        # 获取PDF文件名（不带路径）
        # self.file_name = os.path.basename(in_pdf_path)

        # self.metadata = self.__pdf_reader.getXmpMetadata()

        self.doc_info = self.__pdf_reader.getDocumentInfo()
        #
        self.pages_num = self.__pdf_reader.getNumPages()

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
        elif name_parts[1] == 'json':
            self.bookmark_tree.load_from_json(input)
        else:
            raise Exception(f'Invalid value: {input}')

    def bookmark_tree_to_text_file(self, out_bookmark_path, encoding='utf-8'):
        name_parts = os.path.splitext(out_bookmark_path)

        if name_parts[1].lower() == '.json':
            bookmark_txt = self.bookmark_tree.convert_to_json()
        else:
            bookmark_txt = self.bookmark_tree.convert_to_txt()

        PublicFunc.write_text_file(bookmark_txt, out_bookmark_path, encoding)

    def copy_meta_data(self):
        if not self.doc_info:
            return

        args = {}
        for key, value in list(self.doc_info.items()):
            try:
                args[NameObject(key)] = createStringObject(value)
            except TypeError:
                pass
        self.__pdf_writer.addMetadata(args)

    def remove_bookmarks(self):
        self.__pdf_writer = PdfFileMerger()
        self.__pdf_writer.append(self.__pdf_reader, import_bookmarks=False)

        # copy/preserve existing document info
        self.copy_meta_data()

    def add_bookmarks_to_pdf(self):

        # Another Way: merge `pdf_in` into `pdf_out`, using PyPDF2.PdfFileMerger()
        # self.__pdf_writer = PdfFileMerger()
        # self.__pdf_writer.append(self.__pdf_reader, import_bookmarks=False)

        # 可写的PDF对象，根据不同的模式进行初始化
        self.__pdf_writer = PdfFileWriter()

        if self.mode == Constant.PDF_COPY:
            self.__pdf_writer.cloneDocumentFromReader(self.__pdf_reader)

        elif self.mode == Constant.PDF_NEWLY:
            # self.remove_bookmarks()
            for idx in range(self.pages_num):
                page = self.__pdf_reader.getPage(idx)
                self.__pdf_writer.insertPage(page, idx)

            # copy/preserve existing document info
            self.copy_meta_data()

        self.bookmark_tree.add_to_pdf(self.__pdf_writer)

    def write_to_pdf(self, out_pdf_path):
        # write all data to the given file
        # self.__pdf_writer.write(out_pdf_path)
        # self.__pdf_writer.close()

        # Way2: 保存修改后的PDF文件内容到文件中
        with open(out_pdf_path, 'wb') as fout:
            self.__pdf_writer.write(fout)

    @staticmethod
    def format_bookmark_file(input_bmk_path,
                             output_bmk_path,
                             in_encoding='utf-8',
                             out_encoding='utf-8'):

        # 读取书签文件, 每行为列表的一个元素
        bmk_txt_lines, _ = PublicFunc.read_text_file(input_bmk_path, in_encoding)

        list_reg_patern = [
            # 一级标题：第x章
            (r'^(%s|\s)*(第\d{1,}章)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),

            # 一级标题：第x章
            (r'^(%s|\s)*(第[一二三四五六七八九十〇IV]*章)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),

            # 一级标题：1标题  或  1. 标题
            (r'^(%s|\s)*(\d{1,}\.?)\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             r'\2 '),

            # 二级标题：1.1标题
            (r'^(%s|\s)*(\d{1,}\.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL + r'\2 '),

            # 二级标题：第x节
            (r'^(%s|\s)*(第[一二三四五六七八九十IV]*节)\s*(?=[^.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL + r'\2 '),

            # 三级标题 1.1.1
            (r'^(%s|\s)*(\d{1,}\.\d{1,}\.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL*2 + r'\2 '),

            # 四级标题 1.1.1.1
            (r'^(%s|\s)*(\d{1,}\.\d{1,}\.\d{1,}.\d{1,})\s*(?=[^\d.])' % Constant.MARK_LEVEL_RE,
             Constant.MARK_LEVEL*3 + r'\2 '),

            # 不以数字开头的行，例如：前言
            (r'^(%s|\s)*([^\d%s第])' % (Constant.MARK_LEVEL_RE, Constant.MARK_LEVEL_RE),
             Constant.MARK_LEVEL*0 + r'\2'),

            # 标题与页码间
            (r'(%s|\s)*(-*\d{1,}\r?$)' % Constant.MARK_PAGE_RE,
             Constant.MARK_PAGE + r'\2'),    # 页码
        ]

        res_txt = ''.join(bmk_txt_lines)
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
        
        args.i = 'E:/浏览器下载/基于MPC与预瞄理论的自动驾驶车辆轨迹跟随控制研究_马瀚森.pdf'
        args.o = 'E:/浏览器下载/基于MPC与预瞄理论的自动驾驶车辆轨迹跟随控制研究_马瀚森1.pdf'
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
        exit(2)

    if not os.path.exists(args.i):
        print(f'ERROR: Input file not exist: {args.i}')
        exit(2)

    input_path_parts = os.path.splitext(args.i)

    if ((args.mode in [Constant.ADD, Constant.REMOVE, Constant.EXPORT])
            and input_path_parts[1].lower() != '.pdf'):
        print(f'In mode "{args.mode}", Input file must be PDF format!')
        exit(2)

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
            exit(1)

    # veryfy bookmark path
    if args.mode == Constant.ADD:
        if (not args.bmk):
            #args.bmk = input('input_bookmark_file_path: ')
            args.bmk = input_path_parts[0] + '.txt'

        if not os.path.exists(args.bmk):
            print(f'ERROR: Bookmark file not exist: {args.bmk} ')
            exit(2)


if __name__ == "__main__":

    args = get_cmd_args()
    
    print('\n' + '-'*30)
    print(f'Process mode: \t{args.mode}')
    print(f'Input file: \t{args.i}')
    if args.mode == Constant.ADD:
        print(f'Bookmark file: \t{args.bmk}')
    print(f'Output file: \t{args.o}')

    if args.mode == Constant.ADD:
        pdf_handler = MyPDFHandler(args.i, mode=Constant.PDF_NEWLY)
        print("Read origin pdf file success...")

        pdf_handler.generate_bookmark_tree(args.bmk)
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
# def getCopyText():
#     import win32clipboard as wc
#     import win32con
#     wc.OpenClipboard()
#     copy_text = wc.GetClipboardData(win32con.CF_TEXT)
#     wc.CloseClipboard()
#     return copy_text

# 从配置文件中读取配置信息
# cf = configparser.ConfigParser()
# cf.read('./info.conf')
# pdf_path = cf.get('info', 'pdf_path')
# bookmark_file_path = cf.get('info', 'bookmark_file_path')
# page_offset = cf.getint('info', 'page_offset')
# new_pdf_file_name = cf.get('info', 'new_pdf_file_name')

# def shell():
#     code.interact(
#         banner="PyPDF Bookmarks Shell, pypdfbm_help() for help",
#         local=globals()
#     )
