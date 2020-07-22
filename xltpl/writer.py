# -*- coding: utf-8 -*-

from __future__ import print_function
import six

from .base import BookBase, SheetBase
from .utils import Env
from .utils import tag_test, xv_test
from .xlnode import Row, Cell, EmptyCell, RichCell, TagCell, XvCell, RichTagCell
from .xlrange import SheetRange
from .xlext import CellExtension, RowExtension, SectionExtension, XvExtension
from .ynext import YnExtension
from .richtexthandler import rich_handler

class SheetWriter(SheetBase):
    pass

class BookWriter(BookBase):

    def __init__(self, fname, debug=False):
        self.debug = debug
        self.load(fname)

    def load(self, fname):
        BookBase.load(self, fname)
        self.prepare_env()
        self.sheet_range_list = []
        for rdsheet in self.rdsheet_list:
            sheet_range = self.get_sheet_range(rdsheet)
            if self.debug:
                print('sheet name: ', rdsheet.name)
                print('ranges specified')
                sheet_range.print()
            sheet_range.split()
            if self.debug:
                print('ranges split')
                sheet_range.print_split()
            sheet_range.rich_handler = rich_handler
            tpl_source = sheet_range.to_tag()
            if self.debug:
                print('template source')
                print(tpl_source)
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_range_list.append((sheet_range, jinja_tpl, rdsheet))

    def prepare_env(self):
        self.jinja_env = Env(extensions=[CellExtension, RowExtension, SectionExtension,
                                         YnExtension, XvExtension])
        self.jinja_env.xlsx = False

    def get_sheet_range(self, sheet):
        sheet_range = SheetRange(min_col=1, min_row=1, max_col=sheet.ncols,
                                 max_row=sheet.nrows, index_base=0)
        for rowx in range(sheet.nrows):
            for colx in range(sheet.ncols):
                note = sheet.cell_note_map.get((rowx, colx))
                if note:
                    comment = note.text
                    if tag_test(comment):
                        sheet_range.parse_tag(comment, rowx, colx)
                try:
                    source_cell = sheet.cell(rowx, colx)
                except:
                    sheet_cell = EmptyCell(rowx, colx)
                    sheet_range.add_cell(sheet_cell)
                    continue
                value = source_cell.value
                cty = source_cell.ctype
                rich_text = self.get_rich_text(sheet, rowx, colx)
                if isinstance(value, six.text_type) and xv_test(value):
                    sheet_cell = XvCell(rowx, colx, value, cty)
                elif isinstance(value, six.text_type) and tag_test(value):
                    font = self.get_font(sheet, rowx, colx)
                    if not rich_text:
                        sheet_cell = TagCell(rowx, colx, value, cty, font)
                    else:
                        sheet_cell = RichTagCell(rowx, colx, value, cty, rich_text, font)
                else:
                    if not rich_text:
                        sheet_cell = Cell(rowx, colx, value, cty)
                    else:
                        sheet_cell = RichCell(rowx, colx, value, cty, rich_text)
                sheet_range.add_cell(sheet_cell)
        return sheet_range

    def render_sheet(self, payload, sheet_name, idx):
        sheet_range, jinja_tpl, rdsheet = self.sheet_range_list[idx]
        sheet_writer = SheetWriter(self, rdsheet, sheet_name)
        self.jinja_env.sheet_pos = sheet_range.get_pos(sheet_writer)
        if self.debug:
            print("range positions")
            self.jinja_env.sheet_pos.print()
        rv = jinja_tpl.render(payload)
        sheet_writer.set_mc_ranges()

    def render_book(self, payloads):
        self.create_workbook()
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
        self.wtbook.set_active_sheet(0)

    def render_book2(self, payloads):
        self.create_workbook()
        for payload in payloads:
            idx = self.get_tpl_idx(payload)
            sheet_name = self.get_sheet_name(payload)
            self.render_sheet(payload['ctx'], sheet_name, idx)
        self.wtbook.set_active_sheet(0)

    def render(self, payload):
        idx = self.get_tpl_idx(payload)
        sheet_name = self.get_sheet_name(payload)
        self.render_sheet(payload, sheet_name, idx)
        self.wtbook.set_active_sheet(0)

    def save(self, fname):
        if self.wtbook is not None:
            stream = open(fname, 'wb')
            self.wtbook.save(stream)
            stream.close()
            del self.wtbook

    def finish(self):
        del self.rdbook
        del self.style_list
        del self.font_map