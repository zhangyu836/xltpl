# -*- coding: utf-8 -*-

import copy
from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.utils import get_column_letter

class SheetBase():

    def __init__(self, bookwriter, rdsheet, sheet_name):
        self.workbook = bookwriter.workbook
        self.rdsheet = rdsheet
        self.wtsheet = self.workbook.create_sheet(title=sheet_name)
        self.copy_sheet_settings()
        self.wtsheet.mc_ranges = {}
        self.wtrows = set()
        self.wtcols = set()

    def copy_sheet_settings(self):
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
        self.wtsheet.HeaderFooter = copy.copy(self.rdsheet.HeaderFooter)

    def copy_row_dimension(self, rdrowx, wtrowx):
        if wtrowx in self.wtrows:
            return
        dim = self.rdsheet.row_dimensions.get(rdrowx)
        if dim:
            self.wtsheet.row_dimensions[wtrowx] = copy.copy(dim)
            self.wtsheet.row_dimensions[wtrowx].worksheet = self.wtsheet
            self.wtrows.add(wtrowx)

    def copy_col_dimension(self, rdcolx, wtcolx):
        if wtcolx in self.wtcols:
            return
        rdkey = get_column_letter(rdcolx)
        rddim = self.rdsheet.column_dimensions.get(rdkey)
        if not rddim:
            return
        wtdim = copy.copy(rddim)
        if rdcolx != wtcolx:
            wtkey = get_column_letter(wtcolx)
            wtdim.index = wtkey
            d = wtcolx - rdcolx
            wtdim.min += d
            wtdim.max += d
        else:
            wtkey = rdkey
        self.wtsheet.column_dimensions[wtkey] = wtdim
        self.wtsheet.column_dimensions[wtkey].worksheet = self.wtsheet
        self.wtcols.add(wtcolx)

    def _get_cell(self, rdrowx, rdcolx):
        try:
            return self.rdsheet._cells[(rdrowx, rdcolx)]
        except:
            return Cell(self, row=rdrowx, column=rdcolx)

    def _cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value=None, data_type=None):
        source_cell = self._get_cell(rdrowx, rdcolx)
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if value is None:
            target_cell.value = source_cell._value
            target_cell.data_type = source_cell.data_type
        elif data_type:
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

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx, value=None, data_type=None):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)
        self.merge_cell(rdrowx, rdcolx, wtrowx, wtcolx)
        self._cell(rdrowx, rdcolx, wtrowx, wtcolx, value, data_type)

    def merge_cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
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

    def _mcell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        source_cell = self._get_cell(rdrowx, rdcolx)
        target_cell = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)

    def merge_mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        rdcoords2d = (rdrowx, rdcolx)
        if rdcoords2d in self.rdsheet.mc_top_left_map:
            rlo, rhi, clo, chi = self.wtsheet.mc_ranges.get(rdcoords2d)
            self.wtsheet.mc_ranges[rdcoords2d] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))
        else:
            mc_top_left = self.rdsheet.mc_already_set.get(rdcoords2d)
            if mc_top_left:
                rlo, rhi, clo, chi = self.wtsheet.mc_ranges.get(mc_top_left)
                self.wtsheet.mc_ranges[mc_top_left] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))
            else:
                key = (wt_top, rdcoords2d)
                singel_cell_cr = self.wtsheet.mc_ranges.get(key)
                if singel_cell_cr:
                    rlo, rhi, clo, chi = self.wtsheet.mc_ranges.get(key)
                else:
                    rlo, rhi, clo, chi = wt_top, wtrowx, wtcolx, wtcolx
                self.wtsheet.mc_ranges[key] = (rlo, max(rhi, wtrowx), clo, max(chi, wtcolx))


    def mcell(self, rdrowx, rdcolx, wtrowx, wtcolx, wt_top):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)
        self.merge_mcell(rdrowx, rdcolx, wtrowx, wtcolx, wt_top)
        self._mcell(rdrowx, rdcolx, wtrowx, wtcolx)

    def set_mc_ranges(self):
        for key,crange in self.wtsheet.mc_ranges.items():
            rlo, rhi, clo, chi = crange
            cr = CellRange(min_row=rlo, max_row=rhi, min_col=clo, max_col=chi)
            self.wtsheet.merged_cells.add(cr)

class BookBase():

    def load(self, fname):
        self.workbook = load_workbook(fname)
        self.font_map = {}
        self.sheet_name_map = {}
        self.rdsheet_list = []
        for rdsheet in self.workbook.worksheets:
            self.get_sheet_mc_map(rdsheet)
            self.sheet_name_map[rdsheet.title] = len(self.sheet_name_map)
            self.rdsheet_list.append(rdsheet)
            self.workbook.remove(rdsheet)

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
