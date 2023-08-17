
class Box(object):

    def __init__(self, top, left):
        self.reset_pos(top, left)

    def next_row(self):
        self.bottom += 1
        self.right = self.left

    def next_cell(self):
        self.right += 1

    def reset_pos(self, top, left):
        self.top = top
        self.bottom = top
        self.left = left
        self.right = left

class SheetMixin(object):

    def copy_dimensions(self, rdrowx, rdcolx, wtrowx, wtcolx):
        self.copy_row_dimension(rdrowx, wtrowx)
        self.copy_col_dimension(rdcolx, wtcolx)

    def reset_pos(self, top_left):
        if isinstance(top_left, (tuple, list)):
            self.box = Box(top_left[0], top_left[1])
        elif isinstance(top_left, dict):
            self.box = Box(top_left['top'], top_left['left'])
        else:
            self.box = Box(top_left.top, top_left.left)

    def set_sheet_resource(self, sheet_resource):
        self.merger = sheet_resource.merger
        self.rdsheet = sheet_resource.rdsheet

    def write_row(self, row_node):
        self.box.next_row()

    def write_cell(self, cell_node, rv, cty):
        self.box.next_cell()
        self.merger.merge_cell(cell_node.rowx, cell_node.colx, self.box.bottom, self.box.right)
        self.copy_dimensions(cell_node.rowx, cell_node.colx, self.box.bottom, self.box.right)
        if cell_node.sheet_cell:
            cell_context = self.get_cell_context(cell_node, rv, cty)
            cell_context.finish()
            target_cell = cell_context.target_cell
            if target_cell and hasattr(cell_node, 'ops') and cell_node.ops:
                for (func, func_args) in cell_node.ops:
                    func(*func_args, cell_context)
                cell_node.ops.clear()

    def set_image_ref(self, image_ref):
        image_ref.wtrowx = self.box.bottom
        image_ref.wtcolx = self.box.right + 1
        self.merger.set_image_ref(image_ref)


class BookMixin(object):

    def load(self, fname):
        pass

    def build(self, sheet, index):
        pass

    def add_filter(self, key, value):
        self.jinja_env.filters[key] = value

    def add_global(self, key, value):
        self.jinja_env.globals[key] = value

    def set_jinja_filters(self, **kwargs):
        self.jinja_env.filters.update(kwargs)

    def set_jinja_globals(self, **kwargs):
        self.jinja_env.globals.update(kwargs)

    def get_sheet_writer(self, sheet_resource, sheet_name):
        sheet_writer = self.sheet_writer_map.get(sheet_name)
        if not sheet_writer:
            sheet_writer = self.sheet_writer_cls(self, sheet_resource, sheet_name)
            self.sheet_writer_map[sheet_name] = sheet_writer
        else:
            sheet_writer.set_sheet_resource(sheet_resource)
        return sheet_writer

    def get_sheet_name(self, payload):
        sheet_name = payload.get('sheet_name')
        if sheet_name:
            return sheet_name
        for i in range(9999):
            sheet_name = "sheet%d" % i
            if not self.sheet_writer_map.get(sheet_name):
                return sheet_name
        return "XLSheet"

    def get_sheet_resource(self, payload):
        return self.sheet_resource_map.get_sheet_resource(payload)

    def render_sheet(self, payload, top_left=None):
        sheet_name = self.get_sheet_name(payload)
        sheet_resource = self.get_sheet_resource(payload)
        sheet_writer = self.get_sheet_writer(sheet_resource, sheet_name)
        if top_left:
            sheet_writer.reset_pos(top_left)
        sheet_resource.render_sheet(sheet_writer, payload)
        return sheet_writer.box

    def render_sheets(self, payloads):
        for payload in payloads:
            self.render_sheet(payload)

    def render_book(self, payloads):
        return self.render_sheets(payloads)

    def save(self, fname):
        pass