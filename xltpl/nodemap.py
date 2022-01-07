

class NodeMap():

    def __init__(self):
        self.current_node = None
        self.current_key = ''
        self.last_node = None
        self.last_key = None
        self.node_map = {}

    def set_current_node(self, node):
        self.current_node = node
        self.current_key = node.node_key

    def put(self, key, node):
        self.node_map[key] = node

    def find_lca(self, pre, next):
        # find lowest common ancestor
        next_branch = []
        if pre.depth > next.depth:
            for i in range(next.depth, pre.depth):
                pre.exit()
                # print(pre, 'pre up', pre._parent)
                pre = pre._parent

        elif pre.depth < next.depth:
            for i in range(pre.depth, next.depth):
                next_branch.insert(0, next)
                # print(next, 'next up', next._parent)
                next = next._parent
        if pre is next:
            pass
        else:
            pre_parent = pre._parent
            next_parent = next._parent
            while pre_parent != next_parent:
                # print(pre, next, 'up together')
                pre.exit()
                pre = pre_parent
                pre_parent = pre._parent
                next_branch.insert(0, next)
                next = next_parent
                next_parent = next._parent
            pre.exit()
            if pre_parent._children.index(pre) > pre_parent._children.index(next):
                pre_parent.child_reenter()
                next.reenter()
            else:
                next.enter()

        for next in next_branch:
            next.enter()

    def get_node(self, key):
        if key == self.current_key:
            return self.current_node
        else:
            self.last_key = self.current_key
            self.last_node = self.current_node
            self.current_node = self.node_map.get(key)
            self.current_key = key
            self.find_lca(self.last_node, self.current_node)
        return self.current_node

