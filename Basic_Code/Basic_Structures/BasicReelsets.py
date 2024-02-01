from Basic_Code.Basic_Structures.BasicReel import BasicReel
import xml.etree.ElementTree as ET
import json as j


class BasicReelset:
    def __init__(self, reelset_section_index: int = -1):
        self._reelset_section_index = reelset_section_index
        self._name = ""
        self._reels = []  # List of BasicReel
        self._range = []  # List of two numbers
        self._section_index = -1
        self._section_name = ""
        self._betsIndices = []  # List of numbers

        self._isMainCycle = False
        self._isFreeSpin = False
        self._isRespin = False
        self._isStartScreen = False
        self._isFortuneBet = False

    def MakeReelObj(self):
        return BasicReel()

    def ReadReelset_Json(self, data: dict, reelset_section_index: int = -1):
        self._reelset_section_index = reelset_section_index
        self._reels = [self.MakeReelObj() for _ in range(len(data["Reels"]))]
        for i, reel in enumerate(data["Reels"]):
            self._reels[i].ReadReel_Json(reel, i)
        self._name = data["reelName"]
        self._range = data["range"]
        self._section_index = data["section"]
        self._section_name = data["sectionName"]
        self._betsIndices = [int(num) for num in data.get("betsIndices", '10').split()]
        self._isMainCycle = data.get("isMainCycle", False)
        self._isFreeSpin = data.get("isFreeSpin", False)
        self._isRespin = data.get("isRespin", False)
        self._isStartScreen = data.get("isStartScreen", False)
        self._isFortuneBet = data.get("isFortuneBet", False)

    def ReadReelset_Xml(self, data: ET.Element, reelset_section_index: int = -1):
        self._reelset_section_index = reelset_section_index
        self._reels = [self.MakeReelObj() for _ in range(len(data.findall('Reel')))]
        for reel_index, reel in enumerate(data.findall('Reel')):
            self._reels[reel_index].ReadReel_Xml(reel, reel_index)
        self._name = data.attrib['reelName']
        self._range = [int(num) for num in data.attrib['range'].split(' ')]
        self._section_index = int(data.attrib['section'])
        self._section_name = data.attrib.get('sectionName', None)
        if data.attrib.get('betsIndices', False):
            self._betsIndices = [int(num) for num in data.attrib["betsIndices"].split()]
        if data.attrib.get('isMainCycle', False):
            self._isMainCycle = data.attrib["isMainCycle"] == 'true'
        if data.attrib.get('isFreeSpin', False):
            self._isFreeSpin = data.attrib["isFreeSpin"] == 'true'
        if data.attrib.get('isRespin', False):
            self._isRespin = data.attrib["isRespin"] == 'true'
        if data.attrib.get('isStartScreen', False):
            self._isStartScreen = data.attrib["isStartScreen"] == 'true'
        if data.attrib.get('isFortuneBet', False):
            self._isFortuneBet = data.attrib["isFortuneBet"] == 'true'

    def GetReelsetSectionIndex(self):
        return self._reelset_section_index

    def GetReels(self, reel_index: int = -1):
        if reel_index == -1:
            return self._reels
        else:
            return self._reels[reel_index]

    def GetReelName(self):
        return self._name

    def GetRange(self):
        return self._range

    def GetSectionIndex(self):
        return self._section_index

    def GetSectionName(self):
        return self._section_name

    def GetBetsIndices(self):
        return self._betsIndices

    def IsMainCucle(self):
        return self._isMainCycle

    def IsFreeSpin(self):
        return self._isFreeSpin

    def IsRespin(self):
        return self._isRespin

    def IsStartScreen(self):
        return self._isStartScreen

    def IsFottuneBet(self):
        return self._isFortuneBet

    def __len__(self):
        return len(self._reels)


class BasicReelsets:
    def __init__(self):
        self._section_names = []
        self._reelsets = []  # List of Lists of Reelsets: Section: Reelsets
        self._total_sections_weights = dict()

    def MakeReelsetObj(self):
        return BasicReelset()

    def GetSectionNames(self):
        return self._section_names

    def GetReelsetName(self, section_index: int, reelset_index: int):
        return self._reelsets[section_index][reelset_index].GetReelName()

    def GetReelsetRange(self, section_index: int, reelset_index: int):
        return self._reelsets[section_index][reelset_index].GetRange()

    def GetTotalSectionWeight(self, section_index: int):
        return self._total_sections_weights[section_index]

    def GetReelsetSectionIndex(self, section_index: int, reelset_index: int):
        return self._reelsets[section_index][reelset_index].GetReelsetSectionIndex()

    def GetReelsets(self, section_index: int = -1):
        if section_index == -1:
            return self._reelsets
        else:
            return self._reelsets[section_index]

    def __len__(self):
        return len(self._reelsets)

    def ReadReelsets_Json(self, data: dict):
        self._reelsets = [[] for _ in range(len(data))]
        for str_section_index, section_reelsets in data.items():
            total_section_weight = 0
            int_section_index = int(str_section_index)
            self._reelsets[int_section_index] = [self.MakeReelsetObj() for _ in range(len(section_reelsets))]
            for str_reelset_section_index, reelset in section_reelsets.items():
                int_reelset_section_index = int(str_reelset_section_index)
                current_reelset = self._reelsets[int_section_index][int_reelset_section_index]
                current_reelset.ReadReelset_Json(data=reelset, reelset_section_index=int_reelset_section_index)
                left_weight, right_weight = current_reelset.GetRange()
                total_section_weight += right_weight - left_weight + 1
                if int_section_index+1 > len(self._section_names):
                    self._section_names.append(current_reelset.GetSectionName())
            self._total_sections_weights[int_section_index] = total_section_weight

    def ReadReelsets_Xml(self, data: ET.Element):
        reelsets_by_sections = []
        section_index = 0
        while True:
            section_reelsets = data.findall(".//*[@section='"+str(section_index)+"']")
            if len(section_reelsets) == 0:
                break
            reelsets_by_sections.append(section_reelsets)
            section_index += 1
        self._reelsets = [[] for _ in range(len(reelsets_by_sections))]

        for section_index, reelsets in enumerate(reelsets_by_sections):
            total_section_weight = 0
            self._reelsets[section_index] = [self.MakeReelsetObj() for _ in range(len(reelsets))]
            for reelset_section_index, reelset in enumerate(reelsets):
                cur_reelset = self._reelsets[section_index][reelset_section_index]
                cur_reelset.ReadReelset_Xml(data=reelset, reelset_section_index=reelset_section_index)
                left_weight, right_weight = cur_reelset.GetRange()
                total_section_weight += right_weight - left_weight + 1
                if section_index + 1 > len(self._section_names):
                    self._section_names.append(cur_reelset.GetSectionName())
            self._total_sections_weights[section_index] = total_section_weight


if __name__ == "__main__":
    # file = open(r"C:\Users\VitalijSaiganov\PycharmProjects\XLSX_Writer\Data\Json Stats\SSH\86\SSH_usual_86_no_gamble_stats.json")
    # data = j.load(file)
    # r = BasicReelsets()
    # r.ReadReelsetsJson(data["Reelsets"])
    # print(r.GetSectionNames())

    tree = ET.parse(r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\DTW\86\reels_95.xml")
    reelsets = tree.getroot()[0]

    r = BasicReelsets()
    r.ReadReelsets_Xml(reelsets)
    print(r.GetSectionNames())
    print(r.GetReelsets()[0][0].GetReels()[0].GetSymbols())



