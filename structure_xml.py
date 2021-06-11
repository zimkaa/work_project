# coding: cp1251


# from decimal import Decimal
from decimal import *
from loguru import logger
from time import perf_counter

# from lxml import etree

from xml.dom import minidom
import xml.etree.ElementTree as ET


import xlrd
import xlwt

import config


def read_xlrd(name: str, sheet_index: int = 0) -> list:
    """
    Reading the specified xml sheet
    :param name: file name
    :param sheet_index: sheet to read
    :return: data list
    """
    with xlrd.open_workbook(name, formatting_info=True,
                            encoding_override='WINDOWS-1251') as rb:
        sheet = rb.sheet_by_index(int(sheet_index))
        result = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    return result


def is_record(data: str) -> Decimal:
    """
    Rounding up
    :param data: data string
    :return: Decimal data
    """
    if not data:
        data = '0.0'
    return Decimal(data).quantize(Decimal('1.0'))


def create_rows_section_code5(data: list) -> list:
    """
    Create section code 5
    :param data: all data frome excel
    :return: list of row to section xml
    """
    list_row = []
    for record in data[1:]:
        # print(f"---------record---------\n{record}")
        if record[0]:
            s1 = Decimal(record[0]).quantize(Decimal('1'))
            code = Decimal(record[1]).quantize(Decimal('1'))
            # print(f" 'code': {code} 's1': {s1}" )
            row = {"code": str(code), "s1": str(s1), "s2": str(record[2])}
            col_code1 = is_record(record[3])
            col_code2 = is_record(record[4])
            col_code3 = is_record(record[5]).quantize(
                Decimal('1'), rounding=ROUND_HALF_UP).quantize(Decimal('1.0'))
            col_code4 = is_record(record[6])
            list_row.append([row, col_code1, col_code2, col_code3, col_code4])
    return list_row


def create_rows_section_code3(data: list) -> list:
    """
    Create section code 3
    :param data: all data frome excel
    :return: list of row to section xml
    """
    list_row = []
    for record in data[1:]:
        # print(f"---------record---------\n{record}")
        if record[2]:
            # print(f"---------record[2]---------\n{record[2]}")
            row = {"code": str(record[0])}
            col_code1 = record[2]
            list_row.append([row, col_code1])
    return list_row


def find_number_of_branches(records: list) -> list:
    """
    Find the number of "OKPO"
    :param records: data list
    :return: okpo list
    """
    list_okpo = []
    # print(f"records0 {records[0]}")
    # print(f"records1 {records[1]}")
    for record in records[1:]:
        if record[0] not in list_okpo:
            if record[0]:
                # print(f"record0 {record[0]} ", end="")
                okpo = str(int(record[0]))
                # print(f" - okpo {okpo}")
                list_okpo.append(okpo)
    # print(list_okpo)
    return len(list_okpo)


def main():
    start = perf_counter()
    """open xml"""
    tree = ET.parse(config.NAME_XML_FILE)
    root = tree.getroot()

    """read xls"""
    name_xls_file = config.NAME_XLS_FILE

    """Читается 10й лист 'раздел9'"""
    result = read_xlrd(name_xls_file, 10)

    """finde all fils"""
    print(find_number_of_branches(result))

    section5 = root[1][4]

    """delete all section "code='5'" in xml file"""
    for row in section5.findall('row'):
        section5.remove(row)

    """get list with row section 5"""
    list_section5 = create_rows_section_code5(result)

    """create all section "code='5'" in xml file"""
    for section_row in list_section5:
        row = ET.SubElement(section5, "row", attrib=section_row[0])
        ET.SubElement(row, "col", attrib={"code": "1"}).text = str(section_row[1])
        ET.SubElement(row, "col", attrib={"code": "2"}).text = str(section_row[2])
        ET.SubElement(row, "col", attrib={"code": "3"}).text = str(section_row[3])
        ET.SubElement(row, "col", attrib={"code": "4"}).text = str(section_row[4])

    # """create xml"""
    # tree.write('output.xml', encoding="WINDOWS-1251")

    # tree = ET.parse('output.xml')
    ET.indent(tree, space=" ", level=0)
    tree.write('output2.xml', encoding="WINDOWS-1251")

    stop = perf_counter() - start
    logger.info(f"executinf time {stop}")


if __name__ == '__main__':
    main()
