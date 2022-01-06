
from .misc import Env
from .xlext import NodeExtension, SegmentExtension, XvExtension, ImageExtension, ImagexExtension
from .ynext import YnExtension, YnxExtension

jinja_env = Env(extensions=[NodeExtension, SegmentExtension, YnExtension,
                            XvExtension, ImageExtension])
jinja_envx = Env(extensions=[NodeExtension, SegmentExtension, YnxExtension,
                            XvExtension, ImagexExtension])
