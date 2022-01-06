# -*- coding: utf-8 -*-

import xlrd
import xlwt
import six
import datetime
from decimal import Decimal
from openpyxl.utils.datetime import to_excel

from openpyxl.cell.cell import NUMERIC_TYPES, TIME_TYPES, STRING_TYPES
BOOL_TYPE = bool

def get_type(value):
    if isinstance(value, NUMERIC_TYPES):
        dt = xlrd.XL_CELL_NUMBER
    elif isinstance(value, STRING_TYPES):
        dt = xlrd.XL_CELL_TEXT
    elif isinstance(value, TIME_TYPES):
        dt = xlrd.XL_CELL_DATE
        return to_excel(value), dt
    elif isinstance(value, BOOL_TYPE):
        dt = xlrd.XL_CELL_BOOLEAN
    else:
        return str(value), xlrd.XL_CELL_TEXT
    return value, dt

# adapted from xlutils.filter
class SheetBase():

    def create_worksheet(self, rdsheet, wtsheet_name):
        # these checks should really be done by xlwt!
        if not wtsheet_name:
            raise ValueError('Empty sheet name will result in invalid Excel file!')
        l_wtsheet_name = wtsheet_name.lower()
        if l_wtsheet_name in self.wtbook.wtsheet_names:
            raise ValueError('A sheet named %r has already been added!' % l_wtsheet_name)
        self.wtbook.wtsheet_names.add(l_wtsheet_name)
        l_wtsheet_name = len(wtsheet_name)
        if len(wtsheet_name) > 31:
            raise ValueError('Sheet name cannot be more than 31 characters long, '
                             'supplied name was %i characters long!' % l_wtsheet_name)

        self.rdsheet = rdsheet
        self.wtsheet_name = wtsheet_name
        self.wtsheet = wtsheet = self.wtbook.add_sheet(wtsheet_name, cell_overwrite_ok=True)
        wtsheet.mc_ranges = {}

        # default column width: STANDARDWIDTH, DEFCOLWIDTH
        #
        if rdsheet.standardwidth is not None:
            # STANDARDWIDTH is expressed in units of 1/256 of a
            # character-width, but DEFCOLWIDTH is expressed in units of
            # character-width; we lose precision by rounding to
            # the higher whole number of characters.
            #### XXXX TODO: implement STANDARDWIDTH record in xlwt.
            wtsheet.col_default_width = \
                (rdsheet.standardwidth + 255) // 256
        elif rdsheet.defcolwidth is not None:
            wtsheet.col_default_width = rdsheet.defcolwidth
        #
        # WINDOW2
        #
        wtsheet.show_formulas = rdsheet.show_formulas
        wtsheet.show_grid = rdsheet.show_grid_lines
        wtsheet.show_headers = rdsheet.show_sheet_headers
        wtsheet.panes_frozen = rdsheet.panes_are_frozen
        wtsheet.show_zero_values = rdsheet.show_zero_values
        wtsheet.auto_colour_grid = rdsheet.automatic_grid_line_colour
        wtsheet.cols_right_to_left = rdsheet.columns_from_right_to_left
        wtsheet.show_outline = rdsheet.show_outline_symbols
        wtsheet.remove_splits = rdsheet.remove_splits_if_pane_freeze_is_removed
        wtsheet.selected = rdsheet.sheet_selected
        # xlrd doesn't read WINDOW1 records, so we have to make a best
        # guess at which the active sheet should be:
        # (at a guess, only one sheet should be marked as visible)
        """
        if not self.sheet_visible and rdsheet.sheet_visible:
            self.wtbook.active_sheet = self.wtsheet_index
            wtsheet.sheet_visible = 1
        self.wtsheet_index += 1
        """

        wtsheet.page_preview = rdsheet.show_in_page_break_preview
        wtsheet.first_visible_row = rdsheet.first_visible_rowx
        wtsheet.first_visible_col = rdsheet.first_visible_colx
        wtsheet.grid_colour = rdsheet.gridline_colour_index
        wtsheet.preview_magn = rdsheet.cooked_page_break_preview_mag_factor
        wtsheet.normal_magn = rdsheet.cooked_normal_view_mag_factor
        #
        # DEFAULTROWHEIGHT
        #
        if rdsheet.default_row_height is not None:
            wtsheet.row_default_height = rdsheet.default_row_height
        wtsheet.row_default_height_mismatch = rdsheet.default_row_height_mismatch
        wtsheet.row_default_hidden = rdsheet.default_row_hidden
        wtsheet.row_default_space_above = rdsheet.default_additional_space_above
        wtsheet.row_default_space_below = rdsheet.default_additional_space_below
        #
        # BOUNDSHEET
        #
        wtsheet.visibility = rdsheet.visibility

        #
        # PANE
        #
        if rdsheet.has_pane_record:
            wtsheet.split_position_units_are_twips = True
            wtsheet.active_pane = rdsheet.split_active_pane
            wtsheet.horz_split_pos = rdsheet.horz_split_pos
            wtsheet.horz_split_first_visible = rdsheet.horz_split_first_visible
            wtsheet.vert_split_pos = rdsheet.vert_split_pos
            wtsheet.vert_split_first_visible = rdsheet.vert_split_first_visible

        # print settings
        if hasattr(rdsheet, 'print_headers'):
            wtsheet.print_headers = rdsheet.print_headers
            wtsheet.print_grid = rdsheet.print_grid
            wtsheet.vert_page_breaks = rdsheet.vertical_page_breaks
            wtsheet.horz_page_breaks = rdsheet.horizontal_page_breaks
            wtsheet.header_str = rdsheet.header_str
            wtsheet.footer_str = rdsheet.footer_str
            wtsheet.print_centered_vert = rdsheet.print_centered_vert
            wtsheet.print_centered_horz = rdsheet.print_centered_horz
            wtsheet.left_margin = rdsheet.left_margin
            wtsheet.right_margin = rdsheet.right_margin
            wtsheet.top_margin = rdsheet.top_margin
            wtsheet.bottom_margin = rdsheet.bottom_margin
            wtsheet.paper_size_code = rdsheet.paper_size_code
            wtsheet.print_scaling = rdsheet.print_scaling
            wtsheet.start_page_number = rdsheet.start_page_number
            wtsheet.fit_width_to_pages = rdsheet.fit_width_to_pages
            wtsheet.fit_height_to_pages = rdsheet.fit_height_to_pages
            wtsheet.print_in_rows = rdsheet.print_in_rows
            wtsheet.portrait = rdsheet.portrait
            wtsheet.print_colour = not rdsheet.print_not_colour
            wtsheet.print_draft = rdsheet.print_draft
            wtsheet.print_notes = rdsheet.print_notes
            wtsheet.print_notes_at_end = rdsheet.print_notes_at_end
            wtsheet.print_omit_errors = rdsheet.print_omit_errors
            wtsheet.print_hres = rdsheet.print_hres
            wtsheet.print_vres = rdsheet.print_vres
            wtsheet.header_margin = rdsheet.header_margin
            wtsheet.footer_margin = rdsheet.footer_margin
            wtsheet.copies_num = rdsheet.copies_num

    def copy_row_dimension(self, rdrowx, wtrowx):
        """
        This should be called every time processing of a new
        row in the current sheet starts.

        :param rdrowx: the index of the row in the current sheet from which
                 information for the row to be written will be
                 copied.

        :param wtrowx: the index of the row in sheet to be written to which
                 information will be written for the row being read.
        """
        if wtrowx in self.wtrows or rdrowx not in self.rdsheet.rowinfo_map:
            return
        wtrow = self.wtsheet.row(wtrowx)
        # empty rows may not have a rowinfo record
        rdrow = self.rdsheet.rowinfo_map.get(rdrowx)
        if rdrow:
            wtrow.height = rdrow.height
            wtrow.has_default_height = rdrow.has_default_height
            wtrow.height_mismatch = rdrow.height_mismatch
            wtrow.level = rdrow.outline_level
            wtrow.collapse = rdrow.outline_group_starts_ends  # No kiddin'
            wtrow.hidden = rdrow.hidden
            wtrow.space_above = rdrow.additional_space_above
            wtrow.space_below = rdrow.additional_space_below
            if rdrow.has_default_xf_index:
                wtrow.set_style(self.style_list[rdrow.xf_index])
            self.wtrows.add(wtrowx)

    def copy_col_dimension(self, rdcolx, wtcolx):
        if wtcolx not in self.wtcols and rdcolx in self.rdsheet.colinfo_map:
            rdcol = self.rdsheet.colinfo_map[rdcolx]
            wtcol = self.wtsheet.col(wtcolx)
            wtcol.width = rdcol.width
            wtcol.set_style(self.style_list[rdcol.xf_index])
            wtcol.hidden = rdcol.hidden
            wtcol.level = rdcol.outline_level
            wtcol.collapsed = rdcol.collapsed
            self.wtcols.add(wtcolx)

    def _cell(self, source_cell, rdrowx, rdcolx, wtrowx, wtcolx, value=None, cty=None):
        if value is None:
            value = source_cell.value
            cty = source_cell.ctype
        if cty is None:
            value,cty = get_type(value)

        if cty == xlrd.XL_CELL_EMPTY:
            return
        if source_cell.xf_index is not None:
            style = self.style_list[source_cell.xf_index]
        else:
            style = self.style_list[0]

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

    def cell(self, source_cell, rdrowx, rdcolx, wtrowx, wtcolx, value=None, cty=None):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)
        self._cell(source_cell, rdrowx, rdcolx, wtrowx, wtcolx, value, cty)


class BookBase():

    def load_rdbook(self, fname):
        self.rdbook = rdbook = xlrd.open_workbook(fname, formatting_info=True)
        self.style_list = []
        self.font_map = {}
        for rdxf in rdbook.xf_list:
            wtxf = xlwt.Style.XFStyle()
            #
            # number format
            #
            wtxf.num_format_str = rdbook.format_map[rdxf.format_key].format_str
            #
            # font
            #
            wtf = wtxf.font
            rdf = rdbook.font_list[rdxf.font_index]
            wtf.height = rdf.height
            wtf.italic = rdf.italic
            wtf.struck_out = rdf.struck_out
            wtf.outline = rdf.outline
            wtf.shadow = rdf.outline
            wtf.colour_index = rdf.colour_index
            wtf.bold = rdf.bold  #### This attribute is redundant, should be driven by weight
            wtf._weight = rdf.weight  #### Why "private"?
            wtf.escapement = rdf.escapement
            wtf.underline = rdf.underline_type  ####
            # wtf.???? = rdf.underline #### redundant attribute, set on the fly when writing
            wtf.family = rdf.family
            wtf.charset = rdf.character_set
            wtf.name = rdf.name
            self.font_map[rdxf.font_index] = wtf
            #
            # protection
            #
            wtp = wtxf.protection
            rdp = rdxf.protection
            wtp.cell_locked = rdp.cell_locked
            wtp.formula_hidden = rdp.formula_hidden
            #
            # border(s) (rename ????)
            #
            wtb = wtxf.borders
            rdb = rdxf.border
            wtb.left = rdb.left_line_style
            wtb.right = rdb.right_line_style
            wtb.top = rdb.top_line_style
            wtb.bottom = rdb.bottom_line_style
            wtb.diag = rdb.diag_line_style
            wtb.left_colour = rdb.left_colour_index
            wtb.right_colour = rdb.right_colour_index
            wtb.top_colour = rdb.top_colour_index
            wtb.bottom_colour = rdb.bottom_colour_index
            wtb.diag_colour = rdb.diag_colour_index
            wtb.need_diag1 = rdb.diag_down
            wtb.need_diag2 = rdb.diag_up
            #
            # background / pattern (rename???)
            #
            wtpat = wtxf.pattern
            rdbg = rdxf.background
            wtpat.pattern = rdbg.fill_pattern
            wtpat.pattern_fore_colour = rdbg.pattern_colour_index
            wtpat.pattern_back_colour = rdbg.background_colour_index
            #
            # alignment
            #
            wta = wtxf.alignment
            rda = rdxf.alignment
            wta.horz = rda.hor_align
            wta.vert = rda.vert_align
            wta.dire = rda.text_direction
            # wta.orie # orientation doesn't occur in BIFF8! Superceded by rotation ("rota").
            wta.rota = rda.rotation
            wta.wrap = rda.text_wrapped
            wta.shri = rda.shrink_to_fit
            wta.inde = rda.indent_level
            # wta.merg = ????
            #
            self.style_list.append(wtxf)


    def create_workbook(self):
        self.wtbook = xlwt.Workbook(style_compression=2)
        self.wtbook.dates_1904 = self.rdbook.datemode
        self.wtbook.wtsheet_names = set()

        # Set the default style and the default font
        idx = self.wtbook.add_style(None)
        if idx == 0x10:
            return
        for idx, wtxf in enumerate(self.style_list):
            self.wtbook.add_style(wtxf)
            if idx == 15:
                return
        wtxf = self.style_list[0]
        for _ in range(15 - idx):
            self.wtbook.add_style(wtxf)
        return self.wtbook

    def _get_font(self, index):
        wtf = self.font_map.get(index)
        if wtf:
            return wtf
        else:
            wtf = xlwt.Font()
            rdf = self.rdbook.font_list[index]
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
        xf = self.rdbook.xf_list[sheet.cell_xf_index(rowx, colx)]
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
                xf = self.rdbook.xf_list[sheet.cell_xf_index(rowx, colx)]
                font = self._get_font(xf.font_index)
                rich_text.insert(0, (text, font))
            return rich_text