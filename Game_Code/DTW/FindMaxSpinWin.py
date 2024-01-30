from Basic_Code.Basic_Structures.BasicReelsets import BasicReelsets
from Basic_Code.Basic_Structures.BasicReelsets import BasicReelset
from Basic_Code.Basic_Structures.BasicReel import BasicReel
import xml.etree.ElementTree as ET


class ReelsetCase:
    def __init__(self):
        self.reelset_index = 0
        self.reelset_range = []
        self.reel_indexes = []
        self.reel_cumulative_sum = []

class ReelCase:
    def __init__(self, reelset_index: int, reel_index: int, reel_position: int, cumul_weight: int):
        self.reelset_index = reelset_index
        self.reel_index = reel_index
        self.reel_position = reel_position
        self.cumulative_weight = cumul_weight


BOARD_HEIGHT = 3
BOARD_WIDTH = 5
SYMBOL = 0

def one_condition(window: list):
    if window.count(SYMBOL) >= 2:
        return True
    return False


def FirstReelCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if one_condition(window):
        return True
    else:
        return False


def SecondReelCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if one_condition(window):
        return True
    else:
        return False


def ThirdReelCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if one_condition(window):
        return True
    else:
        return False


def FourthReelCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if one_condition(window):
        return True
    else:
        return False


def FifthReelCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if one_condition(window):
        return True
    else:
        return False

def SHFRCondition(window: list, window_weight: int):
    if window_weight == 0:
        return False
    if window.count(SYMBOL) >= 3:
        return True
    else:
        return False

shfr_conditions = [SHFRCondition] * 20

conditions = [FirstReelCondition,
              SecondReelCondition,
              ThirdReelCondition,
              FourthReelCondition,
              FifthReelCondition]
tree = ET.parse(r"D:\Games Files\#1 SPBR DIGITAL\#17 Nothing But Net\SIMULATIONS\NBN Usual 93\reels_usual.xml")
reelsets = tree.getroot()[0]

reelsets_obj = BasicReelsets()
reelsets_obj.ReadReelsets_Xml(reelsets)
base_reelsets = reelsets_obj.GetReelsets(0)

cases = []
res = 0
for reelset_index, reelset in enumerate(base_reelsets):
    reel_cases_by_reel = [[] for _ in range(len(reelset))]
    for reel_index, reel in enumerate(reelset.GetReels()):
        for reel_position in range(reel.Length()):
            window = reel.TakeWindow(reel_position, BOARD_HEIGHT)
            window_weight = reel.GetWeights()[reel_position]
            cumul_weight = reel.GetCumulWeight(reel_position)

            check_condition = conditions[reel_index](window, window_weight)
            # check_condition = shfr_conditions[reel_index](window, window_weight)

            case = ReelCase(reelset_index, reel_index, reel_position, cumul_weight)
            if check_condition:
                reel_cases_by_reel[reel_index].append(case)

    reel_probs = [0 for _ in range(len(reel_cases_by_reel))]


    for reel_index, reel_cases in enumerate(reel_cases_by_reel):
        for case in reel_cases:
            reel_position = case.reel_position
            reel_weight = sum(reelset.GetReels(reel_index).GetWeights())
            reel_probs[reel_index] += reelset.GetReels(reel_index).GetWeights()[reel_position] / reel_weight

    reelset_prob = 1
    for num in reel_probs:
        reelset_prob *= num
    reelset_range = reelset.GetRange()
    reelset_weight = reelset_range[1] - reelset_range[0] + 1
    section_weight = reelsets_obj.GetTotalSectionWeight(0)
    reelset_prob *= (reelset_weight / section_weight)
    res += reelset_prob
    print(reel_probs)
print()
print('{:,.3f}'.format(1/res))

