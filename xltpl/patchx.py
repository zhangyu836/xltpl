# -*- coding: utf-8 -*-

from openpyxl.writer.excel import ExcelWriter

class ExWriter(ExcelWriter):

    def _write_images(self):
        _written = {}
        for img in self._images:
            if hasattr(img, 'key'):
                key = img.key
                _w = _written.get(key)
                if _w:
                    continue
                else:
                    _written[key] = True
            else:
                print(img, 'not Img')
            self._archive.writestr(img.path[1:], img._data())

ExcelWriter._write_images = ExWriter._write_images

"""
fix  image error
"""
from openpyxl.xml.constants import DRAWING_NS
from openpyxl.drawing.geometry import GeomGuideList
GeomGuideList.tagname = "avLst"
GeomGuideList.namespace = DRAWING_NS
