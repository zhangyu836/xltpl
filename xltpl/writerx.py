# -*- coding: utf-8 -*-

import copy
from jinja2 import Environment

from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

from .utils import tag_test, parse_tag, xv_test
from .xlnode import SheetNodes, Row, Cell, EmptyCell, RichCell, TagCell, XvCell, RichTagCell, SheetPos
from .xlext import CellExtension, RowExtension, SectionExtension, XvExtension
from .ynext import YnExtension
from .richtexthandler import rich_handlerx

class SheetWriter():

    def __init__(self, workbook, rdsheet, sheet_name):
        self.workbook = workbook
        self.rdsheet = rdsheet
        self.wtsheet = self.workbook.create_sheet(title=sheet_name)
        self.copy_sheet_settings()
        self.wtsheet.mc_ranges = {}

    def copy_sheet_settings(self):
        # copy col dimensions
        source = self.rdsheet.column_dimensions
        target = self.wtsheet.column_dimensions
        for key, dim in source.items():
            target[key] = copy.copy(dim)
            target[key].worksheet = self.wtsheet
            
        self.wtsheet.sheet_format = copy.copy(self.rdsheet.sheet_format)
        self.wtsheet.sheet_properties = copy.copy(self.rdsheet.sheet_properties)
        #self.wtsheet.merged_cells = MultiCellRange()
        # copy print settings
        self.wtsheet.page_setup = copy.copy(self.rdsheet.page_setup)
        self.wtsheet.print_options = copy.copy(self.rdsheet.print_options)
        self.wtsheet._print_rows = copy.copy(self.rdsheet._print_rows)
        self.wtsheet._print_cols = copy.copy(self.rdsheet._print_cols)
        self.wtsheet._print_area = copy.copy(self.rdsheet._print_area)
        self.wtsheet.page_margins = copy.copy(self.rdsheet.page_margins)
        self.wtsheet.protection = copy.copy(self.rdsheet.protection)

    def row(self, rdrowx, wtrowx):
        # copy row dimension
        source = self.rdsheet.row_dimensions
        target = self.wtsheet.row_dimensions
        dim = source.get(rdrowx)
        if dim:
            target[wtrowx] = copy.copy(dim)
            target[wtrowx].worksheet = self.wtsheet

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value, data_type=None):
        source_cell  = self.rdsheet._cells[(rdrowx, rdcolx)]
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if data_type:
            target_cell._value = value
            target_cell.data_type = data_type
        else:
            target_cell.value = value
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
                cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
                self.wtsheet.merged_cells.add(cr)
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

class BookWriter():

    def __init__(self, fname):
        self.sheet_list = []
        self.sheet_name_map = {}
        self.load(fname)

    def load(self, fname):
        self.workbook = load_workbook(fname)
        self.font_map = {}
        self.prepare_env()

        for rdsheet in self.workbook.worksheets:
            self.get_sheet_mc_map(rdsheet)
            self.workbook.remove(rdsheet)
            sheet_nodes = self.get_sheet_nodes(rdsheet)
            sheet_nodes.rich_handler = rich_handlerx
            tpl_source = sheet_nodes.to_tag()
            #print(tpl_source)
            jinja_tpl = self.jinja_env.from_string(tpl_source)
            self.sheet_list.append((sheet_nodes, jinja_tpl, rdsheet))
            self.sheet_name_map[rdsheet.title] = len(self.sheet_name_map)

    def prepare_env(self):
        self.jinja_env = Environment(extensions=[CellExtension, RowExtension, SectionExtension, YnExtension, XvExtension])
        self.jinja_env.xlsx = True

    def get_sheet_mc_map(self, sheet):
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
            font = self.workbook._fonts[fontId]
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
        sheet_nodes, jinja_tpl, rdsheet = self.sheet_list[idx]
        sheet_writer = SheetWriter(self.workbook, rdsheet, sheet_name)
        self.jinja_env.sheet_pos = SheetPos(sheet_writer, sheet_nodes, 1, 1)
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
                name = 'XLSheet%d' % len(self.workbook._sheets)
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

    def save(self, fname):
        self.workbook.save(fname)
        for sheet in self.workbook.worksheets:
            self.workbook.remove(sheet)