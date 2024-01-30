import xml.etree.ElementTree as ET

class BasicReel:
    def __init__(self, reel_index: int = -1):
        self._reel_index = reel_index
        self._symbols = []
        self._weights = []

    def TakeWindow(self, index: int,  window_height: int):
        if index > self.Length() - window_height:
            res = self._symbols[index:]
            res += self._symbols[:window_height - (self.Length() - index)]
            return res
        else:
            return self._symbols[index: index + window_height]

    def Length(self):
        return len(self._symbols)

    def ReadReel_Json(self, data: dict, reel_index: int = -1):
        self._reel_index = reel_index
        self._symbols = data["symbols"]
        self._weights = data["weights"]

    def ReadReel_Xml(self, reel: ET.Element, reel_index: int = -1):
        self._reel_index = reel_index
        self._symbols = [int(num) for num in reel.find('Symbols').text.split(',')]
        self._weights = [int(num) for num in reel.find('Weights').text.split(',')]

    def GetReelIndex(self):
        return self._reel_index

    def GetSymbols(self):
        return self._symbols

    def GetWeights(self):
        return self._weights

    def GetCumulWeight(self, reel_position_index: int):
        return sum(self._weights[:reel_position_index])



if __name__ == "__main__":
    g = BasicReel(0)
    g.ReadJsonReel(dict({"symbols": [1,2,3], "weights": [4,5,6]}))
    k = g.TakeWindow(3,3)

    print(k)



