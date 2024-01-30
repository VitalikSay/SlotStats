import numpy as np
import pandas as pd
import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
from Basic_Code.Utils.BasicXMLHadler import Prettify_XML


class BasicSymbol:
    def __init__(self):
        self.__symbol_id = -1
        self.__symbol_name = ""
        self.__symbol_pay = []
        self.__isScatter = False
        self.__isWild = False

    def ReadSymbol_Json(self, symbol_data: dict):
        self.__symbol_id = symbol_data["Index"]
        self.__symbol_name = symbol_data["Name"]
        self.__symbol_pay = symbol_data["Symbol Pay"]
        self.__isScatter = symbol_data.get("Scatter", False)
        self.__isWild = symbol_data.get("Wild", False)

    def GetSymbolPay(self, winline_len: int = -1):
        if winline_len == -1:
            return self.__symbol_pay
        else:
            return self.__symbol_pay[winline_len - 1]

    def GetSymbolId(self):
        return self.__symbol_id

    def GetSymbolName(self):
        return self.__symbol_name

    def IsScatter(self):
        return self.__isScatter

    def IsWild(self):
        return self.__isWild

    def WriteSymbol_FrontTxt(self):
        tab = ' ' * 4
        small_tab = ' ' * 2
        res_str = ""

        for count_index, pay in enumerate(self.__symbol_pay):
            if pay == 0 and not (self.__isScatter and count_index >= 2):
                continue
            res_str += tab + "{ // " + self.__symbol_name + "\n"
            res_str += tab + small_tab + "SymbolId: " + str(self.__symbol_id) + ",\n"
            res_str += tab + small_tab + "SymbolsCount: " + str(count_index + 1) + ",\n"
            res_str += tab + small_tab + "Amount: {\n"
            res_str += tab * 2 + "Win: " + str(pay) + ",\n"
            res_str += tab * 2 + "FreespinsAmount: " + ("INSERT VALUE MANUALLY,\n" if self.__isScatter else "0,\n")
            res_str += tab * 2 + "RespinsAmount: " + ("INSERT VALUE MANUALLY,\n" if self.__isScatter else "0,\n")
            res_str += tab + small_tab + "},\n"
            res_str += tab + "},\n"
        return res_str

    def WriteSymbol_BackXml(self, winlinePaytable: ET.Element):
        for reel_index, pay in enumerate(self.__symbol_pay):
            if pay == 0:
                continue
            payment = ET.SubElement(winlinePaytable, "Payment")
            payment.attrib["numOfSymbols"] = str(reel_index+1)
            payment.attrib["symbolId"] = str(self.__symbol_id)
            payment.attrib["name"] = self.__symbol_name
            payments = ET.SubElement(payment, "Payments")
            payments.attrib["Pay"] = str(pay)

class BasicPayTable:
    def __init__(self):
        self.__symbols = []  # List of BasicSymbols

    def MakeSymbolObj(self):
        return BasicSymbol()

    def ReadPaytable_Json(self, data: dict):
        for i, item in enumerate(data):
            cur_symb = self.MakeSymbolObj()
            cur_symb.ReadSymbol_Json(item)
            assert i == cur_symb.GetSymbolId()
            self.__symbols.append(cur_symb)

    def GetSymbolCount(self):
        return len(self.__symbols)

    def GetSymbolName(self, symbol_id: int):
        return self.__symbols[symbol_id].GetSymbolName()

    def GetSymbolPay(self, symbol_id: int, winline_len: int):
        return self.__symbols[symbol_id].GetSymbolPay(winline_len)

    def IsScatter(self, symbol_id):
        return self.__symbols[symbol_id].IsScatter()

    def IsWild(self, symbol_id):
        return self.__symbols[symbol_id].IsWild()

    def GetPayingCombinations(self):
        paying_combinations = []
        for symbol in self.__symbols:
            temp_paying_combinations = []
            for index, pay in enumerate(symbol.GetSymbolPay()):
                if pay > 0:
                    temp_paying_combinations.append(index)
            if len(temp_paying_combinations) > len(paying_combinations):
                paying_combinations = temp_paying_combinations
        return paying_combinations

    def WritePayTable_FrontTxt(self, out_file_path: str):
        out_file = open(out_file_path, "w")

        tab = " " * 4
        small_tab = " " * 2
        res_str = "export const PAYTABLE = [\n"
        for symbol_index, symbol in enumerate(self.__symbols):
            res_str += symbol.WriteSymbol_FrontTxt()
        res_str += "];\n"

        print(res_str, file=out_file)
        out_file.close()

    def WritePayTable_BackXml(self, out_file_path: str):
        # DEV PAY STARTEGY:
        # < !-- payStrategy:
        # BIGGEST_WIN_ON_LINE == 0
        # LEFT_RIGHT_WIN == 1
        # LONGEST_LINE == 2 -->
        winlinePaytable = ET.Element("WinLinePayTable")
        winlinePaytable.attrib["payStrategy"] = str(2)
        for symbol_index, symbol in enumerate(self.__symbols):
            symbol.WriteSymbol_BackXml(winlinePaytable)
        winlinePaytable = Prettify_XML(winlinePaytable)
        ET.ElementTree(winlinePaytable).write(out_file_path)


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\GLF\86\GLF_usual_86_no_gamble_stats.json")
    data = json.load(file)

    s = BasicPayTable()
    s.ReadPaytable_Json(data["Symbols"])
    s.WritePayTable_BackXml("pytable_back.xml")

    s.WritePayTable_FrontTxt("paytable_front.txt")



