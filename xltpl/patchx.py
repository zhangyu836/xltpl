# -*- coding: utf-8 -*-

import six
from openpyxl.cell.text import Text

"""
A temporary workaround to support rich text.
"""
class IAMRich(six.text_type):

    def __new__(cls, value, rich):
        obj = six.text_type.__new__(cls, value)
        obj.rich = rich
        return obj

    def replace(self, *args, **kwargs):
        s = super(self.__class__, self).replace(*args, **kwargs)
        if s == self:
            return self
        return IAMRich(s, self.rich)

class RichText(Text):

    @property
    def content(self):
        """
        Text stripped of all formatting
        """
        snippets = []
        if self.plain is not None:
            snippets.append(self.plain)
        isrich = False
        for block in self.formatted:
            if block.t is not None:
                isrich = True
                snippets.append(block.t)
        rv = u"".join(snippets)
        if isrich:
            return IAMRich(rv, self)
        return rv

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace, XML_NS
from openpyxl import LXML
from openpyxl.cell._writer import _set_attributes

def etree_write_cell(xf, worksheet, cell, styled=None):

    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        shared_formula = worksheet.formula_attributes.get(cell.coordinate, {})
        formula = SubElement(el, 'f', shared_formula)
        if value is not None:
            formula.text = value[1:]
            value = None

    if cell.data_type == 's':
        if hasattr(value, 'rich') and value.rich:
            sub = value.rich.to_tree(tagname='is')
            el.append(sub)
        else:
            inline_string = SubElement(el, 'is')
            text = SubElement(inline_string, 't')
            text.text = value
            whitespace(text)

    else:
        cell_content = SubElement(el, 'v')
        if value is not None:
            cell_content.text = safe_string(value)

    xf.write(el)


def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(cell.coordinate, {})
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            if hasattr(value, 'rich') and value.rich:
                el = value.rich.to_tree(tagname='is')
                xf.write(el)
            else:
                with xf.element("is"):
                    attrs = {}
                    if value != value.strip():
                        attrs["{%s}space" % XML_NS] = "preserve"
                    el = Element("t", attrs)  # lxml can't handle xml-ns
                    el.text = value
                    xf.write(el)
                    # with xf.element("t", attrs):
                    # xf.write(value)
        else:
            with xf.element("v"):
                if value is not None:
                    xf.write(safe_string(value))

Text.content = RichText.content
if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell

from openpyxl.worksheet import _writer
_writer.write_cell = write_cell

"""
fix  image error
"""
from openpyxl.xml.constants import DRAWING_NS
from openpyxl.drawing.geometry import GeomGuideList
GeomGuideList.tagname = "avLst"
GeomGuideList.namespace = DRAWING_NS
