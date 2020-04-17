# -*- coding: utf-8 -*-

import six
from jinja2 import Environment

from .writerbase import WriterBase, SheetWriter

from .utils import tag_test, parse_tag, xv_test
from .xlnode import SheetNodes, Row, Cell, EmptyCell, RichCell, TagCell, XvCell, RichTagCell, SheetPos
from .xlext import CellExtension, RowExtension, SectionExtension, XvExtension
from .ynext import YnExtension
from .richtexthandler import rich_handler




class BookWriter(WriterBase):

    def __init__(self, fname):
        self.sheet_list = []
        self.sheet_name_map = {}
        self.load(fname)

    def load(self, fname):
        WriterBase.load(self, fname)
        self.prepare_env()
        for rdsheet in self.rdbook.sheets():
            self.get_sheet_mc_map(rdsheet)
            sheet_nodes = self.get_sheet_nodes(rdsheet)
            sheet_nodes.rich_handler = rich_handler
            tpl_source = sheet_nodes.to_tag()
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_list.append((sheet_nodes, jinja_tpl, rdsheet))
            self.sheet_name_map[rdsheet.name] = len(self.sheet_name_map)

    def prepare_env(self):
        self.jinja_env = Environment(extensions=[CellExtension, RowExtension, SectionExtension, YnExtension, XvExtension])
        self.jinja_env.xlsx = False


    def get_rich_text(self, sheet, rowx, colx):
        cell_value = sheet.cell_value(rowx, colx)
        if not cell_value:
            return
        runlist = sheet.rich_text_runlist_map.get((rowx, colx))
        if runlist:
            rich_text = []
            for idx,(start,font_idx) in enumerate(runlist):
                end = None
                if idx != len(runlist) - 1:
                    end = runlist[idx + 1][0]
                text = cell_value[start:end]
                font = self._get_font(font_idx)
                rich_text.append((text, font))
            if runlist[0][0] != 0:
                text = cell_value[:runlist[0][0]]
                xf = self.rdbook.xf_list[sheet.cell_xf_index(rowx, colx)]
                font = self._get_font(xf.font_index)
                rich_text.insert(0, (text, font))
            return rich_text

    def get_sheet_nodes(self, sheet):
        sheet_nodes = SheetNodes()
        for rowx in range(sheet.nrows):
            for colx in range(sheet.row_len(rowx)):
                note = sheet.cell_note_map.get((rowx, colx))
                tag_map = {}
                if note and tag_test(note.text):
                    tag_map = parse_tag(note.text)
                if colx == 0:
                    beforerow = tag_map.get('beforerow')
                    if beforerow:
                        sheet_nodes.add_child(beforerow)
                    sheet_row = Row(rowx)
                    sheet_nodes.add_child(sheet_row)
                beforecell = tag_map.get('beforecell')
                if beforecell:
                    sheet_nodes.add_child(beforecell)

                value = sheet.cell_value(rowx, colx)
                cty = sheet.cell_type(rowx, colx)
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
                sheet_nodes.add_child(sheet_cell)
                aftercell = tag_map.get('aftercell')
                if aftercell:
                    sheet_nodes.add_child(aftercell)
        return sheet_nodes

    def render_sheet(self, payload, sheet_name, idx):
        sheet_nodes, jinja_tpl, rdsheet = self.sheet_list[idx]
        sheet_writer = SheetWriter(self, self.wtbook, rdsheet, sheet_name)
        self.jinja_env.sheet_pos = SheetPos(sheet_writer, sheet_nodes, 0, 0)
        rv = jinja_tpl.render(payload)
        sheet_writer.set_mc_ranges()

    def get_tpl_idx(self, payload):
        idx = payload.get('tpl_idx')
        if not idx:
            name = payload.get('tpl_name')
            if name:
                idx = self.sheet_name_map[name]
            else:
                idx = 0
        return idx

    def get_sheet_name(self, payload, key=None):
        name = payload.get('sheet_name')
        if not name:
            if key:
                name = key
            else:
                name = 'XLSheet%d' % len(self.wtbook.wtsheet_names)
        return name

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


