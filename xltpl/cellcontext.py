import copy
import xlrd
import xlwt
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

class Base(object):

    def __init__(self, sheet_writer, cell_node, value, data_type):
        self.sheet_writer = sheet_writer
        self.cell_node = cell_node
        self.value = value
        self.data_type = data_type

    @property
    def rdsheet(self):
        return self.sheet_writer.rdsheet

    @property
    def wtsheet(self):
        return self.sheet_writer.wtsheet

    @property
    def source_cell(self):
        return self.cell_node.sheet_cell

    @property
    def target_cell(self):
        return None

    @property
    def rdcolx(self):
        return self.cell_node.colx

    @property
    def rdrowx(self):
        return self.cell_node.rowx

    @property
    def wtcolx(self):
        return self.sheet_writer.box.right

    @property
    def wtrowx(self):
        return self.sheet_writer.box.bottom

    def apply_filters(self):
        if hasattr(self.cell_node, 'filters') and self.cell_node.filters:
            for (filter, args) in self.cell_node.filters:
                #print('args', args)
                filter(self, *args)
            self.cell_node.filters.clear()


class CellContextX(Base):

    def __init__(self, sheet_writer, cell_node, value, data_type):
        super().__init__(sheet_writer, cell_node, value, data_type)
        self._target_cell = None

    @property
    def target_cell(self):
        if self._target_cell:
            return self._target_cell
        source = self.source_cell
        wtcolx = self.wtcolx
        wtrowx = self.wtrowx
        value = self.value
        data_type = self.data_type
        target = self.wtsheet.cell(column=wtcolx, row=wtrowx)
        if value is None:
            target._value = source._value
            target.data_type = source.data_type
        elif isinstance(value, STRING_TYPES) and value.startswith('='):
            target.value = value
        elif data_type:
            target._value = value
            target.data_type = data_type
        else:
            target.value = value
        if source.has_style:
            target._style = copy.copy(source._style)
        if source.hyperlink:
            target.hyperlink = copy.copy(source.hyperlink)
        self._target_cell = target
        return self._target_cell

    def get_style(self):
        return self.target_cell.style

    def finish(self):
        self.apply_filters()



class CellContext(Base):

    def __init__(self, sheet_writer, cell_node, value, data_type):
        super().__init__(sheet_writer, cell_node, value, data_type)
        self._style = None

    def get_style(self):
        if self._style:
            return self._style
        source_cell = self.source_cell
        if source_cell.xf_index is not None:
            style = self.sheet_writer.style_list[source_cell.xf_index]
        else:
            style = self.style_list[0]
        self._style = copy.copy(style)
        return self._style

    def set_cell(self):
        cty = self.data_type
        if cty == xlrd.XL_CELL_EMPTY:
            return
        source_cell = self.source_cell
        rdrowx = self.rdrowx
        rdcolx = self.rdcolx
        wtrowx = self.wtrowx
        wtcolx = self.wtcolx
        value = self.value
        style = self.get_style()

        if value is None:
            value = source_cell.value
            cty = source_cell.ctype
        if cty is None:
            value, cty = get_type(value)

        wtrow = self.wtsheet.row(wtrowx)
        if cty == xlrd.XL_CELL_TEXT:
            if isinstance(value, (list, tuple)):
                wtrow.set_cell_rich_text(wtcolx, value, style)
            elif value.startswith('='):
                try:
                    formula = xlwt.Formula(value[1:])
                    wtrow.set_cell_formula(wtcolx, formula, style)
                except BaseException as e:
                    wtrow.set_cell_text(wtcolx, value, style)
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
    def finish(self):
        self.apply_filters()
        self.set_cell()
