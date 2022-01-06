

class SheetMixin(object):

    def reset_pos(self):
        self.current_row_num = self.min_rowx
        self.current_col_num = self.min_colx

    def set_sheet_resource(self, sheet_resource):
        self.merger = sheet_resource.merger
        self.rdsheet = sheet_resource.rdsheet

    def write_row(self, row_node):
        self.current_row_num += 1
        self.current_col_num = self.min_colx

    def write_cell(self, cell_node, rv, cty):
        self.current_col_num += 1
        self.merger.merge_cell(cell_node.rowx, cell_node.colx, self.current_row_num, self.current_col_num)
        if cell_node.sheet_cell:
            self.cell(cell_node.sheet_cell, cell_node.rowx, cell_node.colx, self.current_row_num, self.current_col_num, rv, cty)

    def set_image_ref(self, image_ref, image_key):
        wt_top_left = (self.current_row_num, self.current_col_num+1)
        self.merger.set_image_ref(image_ref, image_key, wt_top_left)


class BookMixin(object):

    def load(self, fname):
        pass

    def build(self, sheet, index):
        pass

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

    def put_sheet_resource(self, index, name, sheet_resource):
        self.sheet_resource_map[index] = sheet_resource
        self.sheet_resource_map[name] = sheet_resource

    def get_sheet_resource(self, payload):
        key = payload.get('tpl_name') or payload.get('tpl_idx') or payload.get('tpl_index')
        return self.sheet_resource_map.get(key) or self.sheet_resource_map.get(0)

    def render_sheet(self, payload):
        sheet_name = self.get_sheet_name(payload)
        sheet_resource = self.get_sheet_resource(payload)
        sheet_writer = self.get_sheet_writer(sheet_resource, sheet_name)
        sheet_resource.render_sheet(sheet_writer, payload)

    def render_sheets(self, payloads):
        for payload in payloads:
            self.render_sheet(payload)

    def render_book(self, payloads):
        return self.render_sheets(payloads)

    def save(self, fname):
        pass