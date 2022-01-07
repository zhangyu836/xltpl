# -*- coding: utf-8 -*-

class TreeProperty(object):

    def __init__(self, name):
        self.name = name
        self._name = '_' + name

    def __set__(self, instance, value):
        instance.__dict__[self._name] = value

    def __get__(self, instance, cls):
        if not hasattr(instance, self._name):
            instance.__dict__[self._name] = getattr(instance._parent, self.name)
        return instance.__dict__[self._name]

class CellTag():

    def __init__(self, cell_tag=dict()):
        self.beforerow = ''
        self.beforecell = ''
        self.aftercell = ''
        self.extracell = ''
        if cell_tag:
            self.__dict__.update(cell_tag)

    def extend(self, other):
        if isinstance(other, CellTag):
            self.beforerow = other.beforerow + self.beforerow
            self.beforecell = other.beforecell + self.beforecell
            self.aftercell += other.aftercell
            self.extracell += other.extracell
