import sys,re
from jinja2 import Environment
from jinja2.exceptions import TemplateSyntaxError
from .xlext import NodeExtension, SegmentExtension, XvExtension, \
    ImageExtension, ImagexExtension, OpExtension, NoopExtension
from .ynext import YnExtension, YnxExtension

class Env(Environment):

    def handle_exception(self, *args, **kwargs):
        exc_type, exc_value, tb = sys.exc_info()
        self.red_fmt = '\033[31m%s\033[0m'
        self.blue_fmt = '\033[34m%s\033[0m'
        self.error_type = self.red_fmt % ('error type:  %s' % exc_type)
        self.error_message = self.red_fmt % ('error message:  %s' % exc_value)
        if exc_type is TemplateSyntaxError:
            lineno = exc_value.lineno
            source = kwargs['source']
            src_lines = source.splitlines()
            self.log_lines(lineno, src_lines)
            self.log_cells(lineno, src_lines)
        Environment.handle_exception(self, *args, **kwargs)

    def get_debug_info(self, line):
        p = re.compile("'(\d*,\d*[,\d]*)'")
        m = p.findall(line)
        debug_info = None
        if len(m) > 0:
            key = m[0]
            node = self.node_map.get_tag_node(key)
            if node:
                debug_info = node.get_debug_info(self.offset)
            else:
                print('---no node---')
        return debug_info

    def log_cells(self, lineno, lines):
        for i, line in enumerate(lines):
            debug_info = self.get_debug_info(line)
            if not debug_info:
                if i + 1 == lineno:
                    print(self.error_message)
                log_str = self.red_fmt % (line)
                print(log_str)
                continue
            if debug_info.value and isinstance(debug_info.value, str):
                line_info = '%s : %s' % (debug_info.address, debug_info.value)
                if i + 1 == lineno:
                    log_str = self.red_fmt % (line_info)
                    print(self.blue_fmt % ('Syntax Error in ' + debug_info.address))
                    print(self.error_message)
                elif i + 1 in [lineno - 1, lineno + 1]:
                    log_str = self.blue_fmt % (line_info)
                else:
                    log_str = line_info
                print(log_str)

    def log_lines(self, lineno, lines):
        for i, line in enumerate(lines):
            debug_info = self.get_debug_info(line)
            if not debug_info:
                if i + 1 == lineno:
                    print(self.error_message)
                line_info = 'line %4d : %s' % (i + 1, line)
                print(self.red_fmt % (line_info))
                continue
            address_line = '   <---   ' + debug_info.address
            line_info = 'line %4d : %s %s' % (i + 1, line, address_line)
            if i + 1 == lineno:
                log_str = self.red_fmt % (line_info)
                print(self.blue_fmt % ('Syntax Error in ' + debug_info.address))
                print(self.error_message)
            elif i + 1 in [lineno - 1, lineno + 1]:
                log_str = self.blue_fmt % (line_info)
            else:
                log_str = line_info
            print(log_str)

    def set_node_map(self, node_map):
        self.node_map = node_map

class JinjaEnv(Env):

    def __init__(self, node_map):
        Env.__init__(self, extensions=[NodeExtension, SegmentExtension, YnExtension,
                                       XvExtension, ImageExtension, NoopExtension])
        self.node_map = node_map
        self.offset = 1


class JinjaEnvx(Env):

    def __init__(self, node_map):
        Env.__init__(self, extensions=[NodeExtension, SegmentExtension, YnxExtension,
                                       XvExtension, ImagexExtension, OpExtension])
        self.node_map = node_map
        self.offset = 0
