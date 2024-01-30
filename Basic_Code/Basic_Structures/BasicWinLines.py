import numpy as np
import pandas as pd
import json
import sys
import xml.etree.ElementTree as ET
from xml.dom import minidom
from Basic_Code.Utils.BasicXMLHadler import Prettify_XML


class BasicWinlines:
    def __init__(self):
        self.__winlines = []

    def ReadWinlines_Json(self, winlines_data: list):
        self.__winlines = np.array(winlines_data)

    def GetWinlines(self, winline_index: int = -1):
        if winline_index == -1:
            return self.__winlines
        else:
            return self.__winlines[winline_index]

    def WriteWinlines_FrontTxt(self, out_file_path: str):
        out_file = open(out_file_path, "w")

        res_str = "export const WIN_LINES =  [\n"
        for winline_index, winline in enumerate(self.__winlines):
            res_str += "    ["
            for num_index, num in enumerate(winline):
                res_str += str(num) + (", " if num_index < len(winline)-1 else "")
            res_str += "]," + "   // index: " + str(winline_index) + "\n"
        res_str += "];\n"

        print(res_str, file=out_file)
        out_file.close()

    def WriteWinlines_BackTxt(self, out_file_path: str):
        out_file = open(out_file_path, "w")

        res_str = "<WinLines>\n"
        for winline_index, winline in enumerate(self.__winlines):
            res_str += '    <Line index="' + str(winline_index)
            res_str += '"    config="'
            for num_ind, num in enumerate(winline):
                res_str += str(num) + (',' if num_ind < len(winline)-1 else '"')
            res_str += " />\n"
        res_str += "</WinLines>\n"

        print(res_str, file=out_file)
        out_file.close()

    def WriteWinlines_BackXml(self, out_file_path: str):
        winlines = ET.Element("WinLines")
        for winline_index, winline in enumerate(self.__winlines):
            txt_winline = ','.join([str(num) for num in winline])
            line = ET.SubElement(winlines, "Line")
            line.attrib["index"] = str(winline_index)
            line.attrib["config"] = txt_winline
        winlines = Prettify_XML(winlines)
        ET.ElementTree(winlines).write(out_file_path)


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\GDD\94\GDD_usual_94_no_gamble_stats.json")
    data = json.load(file)

    f = BasicWinlines()
    f.ReadWinlines_Json(data["Winlines"])

    f.WriteWinlines_FrontTxt("winl.txt")
    f.WriteWinlines_BackXml("winl.xml")
