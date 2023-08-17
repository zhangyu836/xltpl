
def add_filter(filter):
    def wrapper(env, *args):
        cell_node = env.node_map.current_node.current_cell
        cell_node.add_filter(filter, args)
        return ''
    return wrapper


