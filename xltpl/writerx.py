# -*- coding: utf-8 -*-

import copy
from jinja2 import Environment

from .basex import BookBase, SheetBase

from .utils import tag_test, parse_tag, xv_test
from .xlnode import SheetNodes, Row, Cell, EmptyCell, RichCell, TagCell, XvCell, RichTagCell, SheetPos
from .xlext import CellExtension, RowExtension, SectionExtension, XvExtension
from .ynext import YnExtension
from .richtexthandler import rich_handlerx


class SheetWriter(SheetBase):
    pass

class BookWriter(BookBase):

    def __init__(self, fname):
        self.load(fname)

    def load(self, fname):
        BookBase.load(self, fname)
        self.prepare_env()
        self.sheet_nodes_list = []
        for rdsheet in self.rdsheet_list:
            sheet_nodes = self.get_sheet_nodes(rdsheet)
            sheet_nodes.rich_handler = rich_handlerx
            tpl_source = sheet_nodes.to_tag()
            #print(tpl_source)
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_nodes_list.append((sheet_nodes, jinja_tpl, rdsheet))

    def prepare_env(self):
        self.jinja_env = Environment(extensions=[CellExtension, RowExtension, SectionExtension, YnExtension, XvExtension])
        self.jinja_env.xlsx = True

    def get_sheet_nodes(self, sheet):
        sheet_nodes = SheetNodes()

        for rowx in range(1, sheet.max_row + 1):
            for colx in range(1, sheet.max_column + 1):
                source_cell = sheet._cells.get((rowx, colx))
                if not source_cell:
                    if colx == 1:
                        sheet_row = Row(rowx)
                        sheet_nodes.add_child(sheet_row)
                    sheet_cell = EmptyCell(rowx, colx)
                    sheet_nodes.add_child(sheet_cell)
                    continue
                tag_map = {}
                if source_cell.comment:
                    comment = source_cell.comment.text
                    if tag_test(comment):
                        tag_map = parse_tag(comment)

                if colx == 1:
                    beforerow = tag_map.get('beforerow')
                    if beforerow:
                        sheet_nodes.add_child(beforerow)
                    sheet_row = Row(rowx)
                    sheet_nodes.add_child(sheet_row)

                beforecell = tag_map.get('beforecell')
                if beforecell:
                    sheet_nodes.add_child(beforecell)

                value = source_cell._value
                data_type = source_cell.data_type
                rich_text = None
                if hasattr(value, 'rich') and value.rich:
                    rich_text = value.rich
                if data_type == 's' and xv_test(value):
                    sheet_cell = XvCell(rowx, colx, value, data_type)
                elif data_type == 's' and tag_test(value):
                    font = self.get_font(source_cell._style.fontId)
                    if not rich_text:
                        sheet_cell = TagCell(rowx, colx, value, data_type, font)
                    else:
                        sheet_cell = RichTagCell(rowx, colx, value, data_type, rich_text, font)
                else:
                    #if not rich:
                        sheet_cell = Cell(rowx, colx, value, data_type)
                    #else:
                    #    sheet_cell = RichCell(rowx, colx, value, value, data_type)
                sheet_nodes.add_child(sheet_cell)

                aftercell = tag_map.get('aftercell')
                if aftercell:
                    sheet_nodes.add_child(aftercell)
        return sheet_nodes

    def render_sheet(self, payload, sheet_name, idx):
        sheet_nodes, jinja_tpl, rdsheet = self.sheet_nodes_list[idx]
        sheet_writer = SheetWriter(self, rdsheet, sheet_name)
        self.jinja_env.sheet_pos = SheetPos(sheet_writer, sheet_nodes, 1, 1)
        rv = jinja_tpl.render(payload)
        sheet_writer.set_mc_ranges()

    def render_book(self, payloads):
        if isinstance(payloads, dict):
            for key, payload in payloads.items():
                idx = self.get_tpl_idx(payload)
                sheet_name = self.get_sheet_name(payload, key)
                self.render_sheet(payload, sheet_name, idx)
        elif isinstance(payloads, list):
            for payload in payloads:
                idx = self.get_tpl_idx(payload)
                sheet_name = self.get_sheet_name(payload)
                self.render_sheet(payload, sheet_name, idx)

    def render_book2(self, payloads):
        for payload in payloads:
            idx = self.get_tpl_idx(payload)
            sheet_name = self.get_sheet_name(payload)
            self.render_sheet(payload['ctx'], sheet_name, idx)

    def render(self, payload):
        idx = self.get_tpl_idx(payload)
        sheet_name = self.get_sheet_name(payload)
        self.render_sheet(payload, sheet_name, idx)

    def save(self, fname):
        self.workbook.save(fname)
        for sheet in self.workbook.worksheets:
            self.workbook.remove(sheet)