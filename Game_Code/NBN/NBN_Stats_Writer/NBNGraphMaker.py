from Basic_Code.Stats_Writer.BasicGraphMaker import BasicGraphMaker
from Game_Code.NBN.NBN_Structures.NBNSlot import NBNSlot
from Game_Code.NBN.NBN_Structures.NBNStatistics import NBNStatistics
from Game_Code.NBN.NBN_Calculator.NBNStatsCalculator import NBNStatsCalculator


class NBNGraphMaker(BasicGraphMaker):
    def __init__(self, slot: NBNSlot, stats: NBNStatistics, calculation: NBNStatsCalculator):
        super().__init__(slot, stats, calculation)
