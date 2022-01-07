import sys
from jinja2 import Environment
from jinja2.exceptions import TemplateSyntaxError
from .xlext import NodeExtension, SegmentExtension, XvExtension, ImageExtension, ImagexExtension
from .ynext import YnExtension, YnxExtension

class Env(Environment):

    def handle_exception(self, *args, **kwargs):
        exc_type, exc_value, tb = sys.exc_info()
        red_fmt = '\033[31m%s\033[0m'
        blue_fmt = '\033[34m%s\033[0m'
        error_type = red_fmt % ('error type:  %s' % exc_type)
        error_message = red_fmt % ('error message:  %s' % exc_value)
        print(error_type)
        print(error_message)
        if exc_type is TemplateSyntaxError:
            lineno = exc_value.lineno
            source = kwargs['source']
            src_lines = source.splitlines()
            for i, line in enumerate(src_lines):
                if i + 1 == lineno:
                    line_str = red_fmt % ('line %d : %s' % (i + 1, line))
                elif i + 1 in [lineno - 1, lineno + 1]:
                    line_str = blue_fmt % ('line %d : %s' % (i + 1, line))
                else:
                    line_str = 'line %d : %s' % (i + 1, line)
                print(line_str)
        Environment.handle_exception(self, *args, **kwargs)

    def set_node_map(self, node_map):
        self.node_map = node_map

class JinjaEnv(Env):

    def __init__(self, node_map):
        Env.__init__(self, extensions=[NodeExtension, SegmentExtension, YnExtension,
                                       XvExtension, ImageExtension])
        self.node_map = node_map


class JinjaEnvx(Env):

    def __init__(self, node_map):
        Env.__init__(self, extensions=[NodeExtension, SegmentExtension, YnxExtension,
                                       XvExtension, ImagexExtension])
        self.node_map = node_map
