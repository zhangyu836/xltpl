from copy import deepcopy
from openpyxl.drawing.image import Image

# to avoid file closed error
class Cache():

    def __init__(self):
        self.map = {}

    def get_data(self, key):
        return self.map.get(key)

    def set_data(self, key, data):
        self.map[key] = data

    def clear(self):
        self.map.clear()


img_cache = Cache()
data_cache = Cache()

class Img(Image):

    def __init__(self, image):
        self.copy_ref(image)

    def copy_ref(self, image):
        self.anchor = deepcopy(image.anchor)
        self.ref = image.ref
        self.width = image.width
        self.height = image.height
        self.format = image.format

    def set_ref(self, ref):
        super(Img, self).__init__(ref)

    @property
    def path(self):
        key = self.key
        path = img_cache.get_data(key)
        if path:
            return path
        path = super(Img, self).path
        img_cache.set_data(key, path)
        return path

    def _data(self):
        key = self.key
        data = data_cache.get_data(key)
        if data:
            return data
        data = super(Img, self)._data()
        data_cache.set_data(key, data)
        return data

    @property
    def key(self):
        return id(self.ref)




