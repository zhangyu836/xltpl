# -*- coding: utf-8 -*-

import copy
from jinja2 import Environment

from openpyxl import load_workbook
from openpyxl.worksheet.copier import WorksheetCopy
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

from xltpl.utils import tag_test, parse_tag
from xltpl.xlnode import Sheet, Row, Cell, EmptyCell, RichCell, TagCell, RichTagCell
from xltpl.xlext import CellExtension, RowExtension, SectionExtension
from xltpl.ynext import YnExtension
from xltpl.richtexthandler import  RichTextHandlerX


class SheetCopy(WorksheetCopy):

    def copy_sheet(self):
        self.copy_col_dimensions()
        self.target.sheet_format = copy.copy(self.source.sheet_format)
        self.target.sheet_properties = copy.copy(self.source.sheet_properties)
        #self.target.merged_cells = MultiCellRange()
        self.copy_print_settings()

    def copy_col_dimensions(self):
        source = self.source.column_dimensions
        target = self.target.column_dimensions
        for key, dim in source.items():
            target[key] = copy.copy(dim)
            target[key].worksheet = self.target

    def copy_row_dimension(self, rdrowx, wtrowx):
        source = self.source.row_dimensions
        target = self.target.row_dimensions
        dim = source.get(rdrowx)
        if dim:
            target[wtrowx] = copy.copy(dim)
            target[wtrowx].worksheet = self.target

    def copy_print_settings(self):
        self.target.page_setup = copy.copy(self.source.page_setup)
        self.target.print_options = copy.copy(self.source.print_options)
        self.target._print_rows = copy.copy(self.source._print_rows)
        self.target._print_cols = copy.copy(self.source._print_cols)
        self.target._print_area = copy.copy(self.source._print_area)
        self.target.page_margins = copy.copy(self.source.page_margins)
        self.target.protection = copy.copy(self.source.protection)

class XlsxWriter():

    def __init__(self, fname):
        self.sheet_tpl_list = []
        self.sheet_tpl_name_map = {}
        self.rich_text_handler = RichTextHandlerX()
        self.load(fname)

    def load(self, fname):
        self.src_rdbook = load_workbook(fname)
        self.font_map = {}
        self.prepare_env()

        for src_rdsheet in self.src_rdbook.worksheets:
            self.get_sheet_info(src_rdsheet)
            self.src_rdbook.remove(src_rdsheet)
            sheet_tpl = self.parse_sheet(src_rdsheet)
            tpl_source = sheet_tpl.to_tag()
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_tpl_list.append((sheet_tpl, jinja_tpl, src_rdsheet))
            self.sheet_tpl_name_map[src_rdsheet.title] = len(self.sheet_tpl_name_map)

    def prepare_env(self):
        self.jinja_env = Environment(extensions=[CellExtension, RowExtension, SectionExtension, YnExtension])
        self.jinja_env.xlsx = True

    def get_sheet_info(self, sheet):
        mc_map = {}
        mc_nfa = {}
        for crange in sheet.merged_cells.ranges:
            rlo, rhi, clo, chi = crange.min_row, crange.max_row, crange.min_col, crange.max_col
            mc_map[(rlo, clo)] = (rlo, rhi, clo, chi)
            for rowx in range(rlo, rhi + 1):
                for colx in range(clo, chi + 1):
                    mc_nfa[(rowx, colx)] = (rlo, clo)
        sheet.mc_top_left_map = mc_map
        sheet.mc_already_set = mc_nfa

    def get_font(self, fontId):
        ifont = self.font_map.get(fontId)
        if ifont:
            return ifont
        else:
            font = self.src_rdbook._fonts[fontId]
            ifont = InlineFont()
            ifont.rFont = font.name
            ifont.charset = font.charset
            ifont.family = font.family
            ifont.b = font.b
            ifont.i = font.i
            ifont.strike = font.strike
            ifont.outline = font.outline
            ifont.shadow = font.shadow
            ifont.condense = font.condense
            ifont.extend = font.extend
            ifont.color = font.color
            ifont.sz = font.sz
            ifont.u = font.u
            ifont.vertAlign = font.vertAlign
            ifont.scheme = font.scheme
            self.font_map[fontId] = ifont
            return ifont

    def parse_sheet(self, sheet):
        sheet_tpl = Sheet(self, first_row=1, first_col=1)
        sheet_tpl.handler = self.rich_text_handler

        for rowx in range(1, sheet.max_row + 1):
            for colx in range(1, sheet.max_column + 1):
                source_cell = sheet._cells.get((rowx, colx))
                if not source_cell:
                    if colx == 1:
                        sheet_row = Row(rowx, first_row=1)
                        sheet_tpl.add_child(sheet_row)
                    sheet_cell = EmptyCell(rowx, colx)
                    sheet_tpl.add_child(sheet_cell)
                    continue
                tag_map = {}
                if source_cell.comment:
                    comment = source_cell.comment.text
                    if tag_test(comment):
                        tag_map = parse_tag(comment)

                if colx == 1:
                    beforerow = tag_map.get('beforerow')
                    if beforerow:
                        sheet_tpl.add_child(beforerow)
                    sheet_row = Row(rowx, first_row=1)
                    sheet_tpl.add_child(sheet_row)

                beforecell = tag_map.get('beforecell')
                if beforecell:
                    sheet_tpl.add_child(beforecell)

                value = source_cell._value
                data_type = source_cell.data_type
                rich_text = None
                if hasattr(value, 'rich') and value.rich:
                    rich_text = value.rich
                if data_type == 's' and tag_test(value):
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
                name = 'XLSheet%d' % len(self.src_rdbook._sheets)
        return name

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

    def sheet(self, rdsheet, sheet_name):
        wtsheet = self.src_rdbook.create_sheet(title=sheet_name)
        self.copier = SheetCopy(source_worksheet=rdsheet, target_worksheet=wtsheet)
        self.copier.copy_sheet()
        self.rdsheet = rdsheet
        self.wtsheet = wtsheet
        self.wtsheet.mc_ranges = {}
        self.wtsheet.mc_ranges_already_set = {}

    def row(self, rdrowx, wtrowx):
        self.copier.copy_row_dimension(rdrowx, wtrowx)

    def set_cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value, data_type):
        source_cell  = self.rdsheet._cells[(rdrowx, rdcolx)]
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        target_cell._value = value
        target_cell.data_type = data_type
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)
        if source_cell.hyperlink:
            target_cell._hyperlink = copy.copy(source_cell.hyperlink)
        #if source_cell.comment:
        #    target_cell.comment = copy.copy(source_cell.comment)

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

    def set_mc_ranges(self):
        for key,crange in self.wtsheet.mc_ranges.items():
            rlo, rhi, clo, chi = crange
            cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
            self.wtsheet.merged_cells.add(cr)
        for key,crange in self.wtsheet.mc_ranges_already_set.items():
            rlo, rhi, clo, chi = crange
            cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
            self.wtsheet.merged_cells.add(cr)

    def save(self, fname):
        self.src_rdbook.save(fname)
        for sheet in self.src_rdbook.worksheets:
            self.src_rdbook.remove(sheet)