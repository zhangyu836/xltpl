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


