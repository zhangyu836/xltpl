
from .nodemap import node_map
class SheetResource():

    def __init__(self, rdsheet, sheet_tree, jinja_env, merger):
        self.rdsheet = rdsheet
        self.sheet_tree = sheet_tree
        self.tpl = sheet_tree.to_tag()
        #print(self.tpl)
        self.merger = merger
        self.jinja_tree = jinja_env.from_string(self.tpl)


    def render_sheet(self, sheet_writer, payload):
        node_map.set_current_node(self.sheet_tree)
        self.sheet_tree.set_sheet_writer(sheet_writer)
        self.jinja_tree.render(payload)
        self.merger.collect_range(sheet_writer.wtsheet)
