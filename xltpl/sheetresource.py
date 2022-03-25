
class SheetResource():

    def __init__(self, book_writer, rdsheet, index, jinja_env):
        self.rdsheet = rdsheet
        self.merger = book_writer.merger_cls(rdsheet)
        #self.jinja_env = jinja_env
        self.sheet_tree = book_writer.build(rdsheet, index, self.merger)
        self.tpl = self.sheet_tree.to_tag()
        #print(self.tpl)
        self.jinja_tree = jinja_env.from_string(self.tpl)

    def render_sheet(self, sheet_writer, payload):
        self.sheet_tree.set_sheet_writer(sheet_writer)
        self.jinja_tree.render(payload)
        self.merger.collect_range(sheet_writer.wtsheet)

class SheetState() :

    def __init__(self, book_writer, rdsheet, jinja_env):
        self.book_writer = book_writer
        self.rdsheet = rdsheet
        self.jinja_env = jinja_env
        self.sheet_resource = None

    def get_sheet_resource(self):
        if not self.sheet_resource:
            self.sheet_resource = SheetResource(self.book_writer, self.rdsheet, self.index, self.jinja_env)
        return self.sheet_resource


class SheetResourceMap():

    def __init__(self, book_writer, jinja_env):
        self.sheet_state_map = {}
        self.sheet_state_list = []
        self.book_writer = book_writer
        self.jinja_env = jinja_env

    def add(self, rdsheet, name, index):
        sheet_state = SheetState(self.book_writer, rdsheet, self.jinja_env)
        self.sheet_state_map[name] = sheet_state
        self.sheet_state_map[index] = sheet_state
        self.sheet_state_list.append(sheet_state)
        sheet_state.index = index
        sheet_state.name = name

    def get_sheet_resource(self, payload):
        key = payload.get('tpl_name') or payload.get('tpl_idx') or payload.get('tpl_index')
        sheet_state = self.sheet_state_map.get(key) or self.sheet_state_map.get(0)
        return sheet_state.get_sheet_resource()
