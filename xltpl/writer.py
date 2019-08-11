# -*- coding: utf-8 -*-

import copy
import xlrd
import six
from jinja2 import Environment

from xlwt import Font
from xlutils.filter import BaseWriter

from xltpl.utils import tag_test, parse_tag
from xltpl.xlnode import Sheet, Row, Cell, EmptyCell, RichCell, TagCell, RichTagCell
from xltpl.xlext import CellExtension, RowExtension, SectionExtension
from xltpl.ynext import YnExtension
from xltpl.richtexthandler import RichTextHandler

class XlsWriter(BaseWriter, object):

    def __init__(self, fname):
        self.sheet_tpl_list = []
        self.sheet_tpl_name_map = {}
        self.rich_text_handler = RichTextHandler()
        self.load(fname)

    def load(self, fname):
        self.src_rdbook = xlrd.open_workbook(fname, formatting_info=True)
        self.workbook(self.src_rdbook, '')
        if not hasattr(self, 'font_map'):
            self.font_map = {}
        self.prepare_env()
        for src_rdsheet in self.src_rdbook.sheets():
            self.get_sheet_info(src_rdsheet)
            sheet_tpl = self.parse_sheet(src_rdsheet)
            tpl_source = sheet_tpl.to_tag()
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_tpl_list.append((sheet_tpl, jinja_tpl, src_rdsheet))
            self.sheet_tpl_name_map[src_rdsheet.name] = len(self.sheet_tpl_name_map)
        self.wtbook_holder = self.wtbook
        self.wtbook = None

    def prepare_env(self):
        self.jinja_env = Environment(extensions=[CellExtension, RowExtension, SectionExtension, YnExtension])

    def get_sheet_info(self, sheet):
        mc_map = {}
        mc_nfa = {}
        for crange in sheet.merged_cells:
            rlo, rhi, clo, chi = crange
            mc_map[(rlo, clo)] = crange
            for rowx in range(rlo, rhi):
                for colx in range(clo, chi):
                    mc_nfa[(rowx, colx)] = (rlo, clo)
        sheet.mc_top_left_map = mc_map
        sheet.mc_already_set = mc_nfa

    def _get_font(self, index):
        wtf = self.font_map.get(index)
        if wtf:
            return wtf
        else:
            wtf = Font()
            rdf = self.src_rdbook.font_list[index]
            wtf.height = rdf.height
            wtf.italic = rdf.italic
            wtf.struck_out = rdf.struck_out
            wtf.outline = rdf.outline
            wtf.shadow = rdf.outline
            wtf.colour_index = rdf.colour_index
            wtf.bold = rdf.bold
            wtf._weight = rdf.weight
            wtf.escapement = rdf.escapement
            wtf.underline = rdf.underline_type
            wtf.family = rdf.family
            wtf.charset = rdf.character_set
            wtf.name = rdf.name
            self.font_map[index] = wtf
            return wtf

    def get_font(self, sheet, rowx, colx):
        xf = self.src_rdbook.xf_list[sheet.cell_xf_index(rowx, colx)]
        return self._get_font(xf.font_index)

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
                xf = self.src_rdbook.xf_list[sheet.cell_xf_index(rowx, colx)]
                font = self._get_font(xf.font_index)
                rich_text.insert(0, (text, font))
            return rich_text

    def parse_sheet(self, sheet):
        sheet_tpl = Sheet(self)
        sheet_tpl.handler = self.rich_text_handler
        for rowx in range(sheet.nrows):
            for colx in range(sheet.row_len(rowx)):
                note = sheet.cell_note_map.get((rowx, colx))
                tag_map = {}
                if note and tag_test(note.text):
                    tag_map = parse_tag(note.text)
                if colx == 0:
                    beforerow = tag_map.get('beforerow')
                    if beforerow:
                        sheet_tpl.add_child(beforerow)
                    sheet_row = Row(rowx)
                    sheet_tpl.add_child(sheet_row)
                beforecell = tag_map.get('beforecell')
                if beforecell:
                    sheet_tpl.add_child(beforecell)

                value = sheet.cell_value(rowx, colx)
                cty = sheet.cell_type(rowx, colx)
                rich_text = self.get_rich_text(sheet, rowx, colx)
                if isinstance(value, six.text_type) and tag_test(value):
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
                sheet_tpl.add_child(sheet_cell)
                aftercell = tag_map.get('aftercell')
                if aftercell:
                    sheet_tpl.add_child(aftercell)
        return sheet_tpl

    def render_sheet(self, payload, sheet_name, idx):
        sheet_tpl, jinja_tpl, src_rdsheet = self.sheet_tpl_list[idx]
        self.jinja_env.sheet_tpl = sheet_tpl
        self.sheet(src_rdsheet, sheet_name)
        rv = jinja_tpl.render(payload)
        self.set_mc_ranges()
        sheet_tpl.reset()

    def get_tpl_idx(self, payload):
        idx = payload.get('tpl_idx')
        if not idx:
            name = payload.get('tpl_name')
            if name:
                idx = self.sheet_tpl_name_map[name]
            else:
                idx = 0
        return idx

    def get_sheet_name(self, payload, key=None):
        name = payload.get('sheet_name')
        if not name:
            if key:
                name = key
            else:
                name = 'XLSheet%d' % len(self.wtsheet_names)
        return name

    def render_book(self, payloads):
        self.wtbook = copy.deepcopy(self.wtbook_holder)
        self.wtsheet_names = set()
        self.wtsheet_index = 0
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
        self.wtbook = copy.deepcopy(self.wtbook_holder)
        self.wtsheet_names = set()
        self.wtsheet_index = 0
        for payload in payloads:
            idx = self.get_tpl_idx(payload)
            sheet_name = self.get_sheet_name(payload)
            self.render_sheet(payload['ctx'], sheet_name, idx)
        self.wtbook.set_active_sheet(0)


    def render(self, payload):
        self.wtbook = copy.deepcopy(self.wtbook_holder)
        self.wtsheet_names = set()
        self.wtsheet_index = 0
        idx = self.get_tpl_idx(payload)
        sheet_name = self.get_sheet_name(payload)
        self.render_sheet(payload, sheet_name, idx)
        self.wtbook.set_active_sheet(0)


    def sheet(self, rdsheet, wtsheet_name):
        BaseWriter.sheet(self, rdsheet, wtsheet_name)
        self.wtsheet.selected = 0
        self.wtsheet.mc_ranges = {}
        self.wtsheet.mc_ranges_already_set = {}

    def set_cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value, cty):
        cell = self.rdsheet.cell(rdrowx, rdcolx)
        if wtcolx not in self.wtcols and rdcolx in self.rdsheet.colinfo_map:
            rdcol = self.rdsheet.colinfo_map[rdcolx]
            wtcol = self.wtsheet.col(wtcolx)
            wtcol.width = rdcol.width
            wtcol.set_style(self.style_list[rdcol.xf_index])
            wtcol.hidden = rdcol.hidden
            wtcol.level = rdcol.outline_level
            wtcol.collapsed = rdcol.collapsed
            self.wtcols.add(wtcolx)

        if cty == xlrd.XL_CELL_EMPTY:
            return
        if cell.xf_index is not None:
            style = self.style_list[cell.xf_index]
        else:
            style = self.style_list[0]
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.mc_top_left_map:
            if self.wtsheet.mc_ranges.get(rdcoords2d):
                rlo, rhi, clo, chi = self.wtsheet.mc_ranges.get(rdcoords2d)
                self.wtsheet.mc_ranges_already_set[(rlo, clo)] = (rlo, rhi, clo, chi)
            self.wtsheet.mc_ranges[rdcoords2d] = (wtrowx, wtrowx, wtcolx, wtcolx)
        else:
            mc_top_left = self.rdsheet.mc_already_set.get(rdcoords2d)
            if mc_top_left:
                rlo, rhi, clo, chi = self.wtsheet.mc_ranges.get(mc_top_left)
                self.wtsheet.mc_ranges[mc_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))

        wtrow = self.wtsheet.row(wtrowx)
        if cty == xlrd.XL_CELL_TEXT:
            if isinstance(value, (list, tuple)):
                wtrow.set_cell_rich_text(wtcolx, value, style)
            else:
                wtrow.set_cell_text(wtcolx, value, style)
        elif cty == xlrd.XL_CELL_NUMBER or cty == xlrd.XL_CELL_DATE:
            wtrow.set_cell_number(wtcolx, value, style)
        elif cty == xlrd.XL_CELL_BLANK:
            wtrow.set_cell_blank(wtcolx, style)
        elif cty == xlrd.XL_CELL_BOOLEAN:
            wtrow.set_cell_boolean(wtcolx, value, style)
        elif cty == xlrd.XL_CELL_ERROR:
            wtrow.set_cell_error(wtcolx, value, style)
        else:
            raise Exception(
                "Unknown xlrd cell type %r with value %r at (sheet=%r,rowx=%r,colx=%r)" \
                % (cty, value, self.rdsheet.name, rdrowx, rdcolx)
                )

    def set_mc_ranges(self):
        for key,crange in self.wtsheet.mc_ranges.items():
            self.wtsheet.merged_ranges.append(crange)
        for key,crange in self.wtsheet.mc_ranges_already_set.items():
            self.wtsheet.merged_ranges.append(crange)

    def save(self, fname):
        self.wtname = fname
        if self.wtbook is not None:
            stream = open(self.wtname, 'wb')
            self.wtbook.save(stream)
            stream.close()

    def finish(self):
        del self.src_rdbook
        del self.wtbook_holder
        del self.wtbook
        del self.wtsheet
        del self.style_list
        del self.font_map


