# -*- coding: utf-8 -*-

import six

from .base import SheetBase, BookBase
from .writermixin import SheetMixin, BookMixin
from .utils import tag_test
from .xlnode import Tree, Row, Cell, EmptyCell, Node, create_cell
from .jinja import JinjaEnv
from .nodemap import NodeMap
from .sheetresource import SheetResource
from .richtexthandler import rich_handler
from .merger import Merger
from .config import config

class SheetWriter(SheetBase, SheetMixin):

    def __init__(self, bookwriter, sheet_resource, sheet_name):
        self.rdbook = bookwriter.rdbook
        self.wtbook = bookwriter.wtbook
        self.style_list = bookwriter.style_list
        self.merger = sheet_resource.merger
        self.rdsheet = sheet_resource.rdsheet
        self.create_worksheet(self.rdsheet, sheet_name)
        self.wtrows = set()
        self.wtcols = set()
        self.min_rowx = -1
        self.min_colx = -1
        self.reset_pos()

class BookWriter(BookBase, BookMixin):
    sheet_writer_cls = SheetWriter

    def __init__(self, fname, debug=False):
        config.debug = debug
        self.load(fname)

    def load(self, fname):
        self.workbook = self.load_rdbook(fname)
        self.font_map = {}
        self.sheet_writer_map = {}
        self.sheet_resource_map = {}
        self.node_map = NodeMap()
        self.jinja_env = JinjaEnv(self.node_map)
        for index,rdsheet in enumerate(self.rdbook.sheets()):
            sheet_tree = self.build(rdsheet, index)
            merger = Merger(rdsheet)
            sheet_resource = SheetResource(rdsheet, sheet_tree, self.jinja_env, merger)
            self.put_sheet_resource(index, rdsheet.name, sheet_resource)

    def build(self, sheet, index):
        tree = Tree(index, self.node_map)
        for rowx in range(sheet.nrows):
            row_node = Row(rowx)
            tree.add_child(row_node)
            for colx in range(sheet.ncols):
                try:
                    sheet_cell = sheet.cell(rowx, colx)
                except:
                    cell_node = EmptyCell(rowx, colx)
                    tree.add_child(cell_node)
                    continue
                value = sheet_cell.value
                cty = sheet_cell.ctype
                rich_text = self.get_rich_text(sheet, rowx, colx)
                if isinstance(value, six.text_type):
                    if not tag_test(value):
                        if rich_text:
                            cell_node = Cell(sheet_cell, rowx, colx, rich_text, cty)
                        else:
                            cell_node = Cell(sheet_cell, rowx, colx, value, cty)
                    else:
                        font = self.get_font(sheet, rowx, colx)
                        cell_node = create_cell(sheet_cell, rowx, colx, value, rich_text, cty, font, rich_handler)
                else:
                    cell_node = Cell(sheet_cell, rowx, colx, value, cty)
                tree.add_child(cell_node)
            tree.add_child(Node())  #
        return tree

    def render_sheet(self, payload):
        if not hasattr(self, 'wtbook') or self.wtbook is None:
            self.create_workbook()
        BookMixin.render_sheet(self, payload)

    def save(self, fname):
        if self.wtbook is not None:
            stream = open(fname, 'wb')
            self.wtbook.save(stream)
            stream.close()
            del self.wtbook
        self.sheet_writer_map.clear()
