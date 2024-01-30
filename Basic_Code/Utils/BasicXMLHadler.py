import xml.etree.ElementTree as ET
from xml.dom import minidom


def Prettify_XML(root_elem: ET.Element):
    rough_string = ET.tostring(root_elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return ET.fromstring(reparsed.toprettyxml(indent="    "))