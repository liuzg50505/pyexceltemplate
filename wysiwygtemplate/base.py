
class WrapedDict:

    def __init__(self, innerdict) -> None:
        super().__init__()
        self.innerdict = innerdict
        self.dict = {}

    def add(self, key,value):
        self.dict[key] = value

    def addall(self, dictobj):
        for k,v in dictobj.items():
            self.add(k, v)

    def __getitem__(self, item):
        if item in self.dict: return self.dict[item]
        if item in self.innerdict: return self.innerdict[item]
        return None

    def __contains__(self, item):
        if item in self.dict: return True
        if item in self.innerdict: return True
        return False







