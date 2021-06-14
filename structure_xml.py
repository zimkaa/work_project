# coding: cp1251

import decimal
from loguru import logger
from time import perf_counter

from xlrd import open_workbook  # type: ignore
import xml.etree.ElementTree as ElT

import config


def read_xlrd(name: str, sheet_index: int = 0) -> list:
    """
    Reading the specified xml sheet
    :param name: file name
    :param sheet_index: sheet to read
    :return: data list
    """
    with open_workbook(name, formatting_info=True,
                       encoding_override='WINDOWS-1251') as rb:
        sheet = rb.sheet_by_index(int(sheet_index))
        result = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    return result


def is_record(data: str) -> decimal.Decimal:
    """
    Rounding up
    :param data: data string
    :return: Decimal data
    """
    if not data:
        data = '0.0'
    return decimal.Decimal(data).quantize(decimal.Decimal('1.0'))


def create_rows_section_code5(data: list) -> list:
    """
    Create section code 5
    :param data: all data frome excel
    :return: list of row to section xml
    """
    list_row = []
    for record in data[1:]:
        if record[0]:
            s1 = decimal.Decimal(record[0]).quantize(decimal.Decimal('1'))
            code = decimal.Decimal(record[1]).quantize(decimal.Decimal('1'))
            row = {"code": str(code), "s1": str(s1), "s2": str(record[2])}
            col_code1 = is_record(record[3])
            col_code2 = is_record(record[4])
            col_code3 = is_record(record[5]).quantize(
                decimal.Decimal('1'), rounding=decimal.ROUND_HALF_UP)\
                .quantize(decimal.Decimal('1.0'))
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
        if record[2]:
            row = {"code": str(record[0])}
            col_code1 = record[2]
            list_row.append([row, col_code1])
    return list_row


def find_number_of_branches(records: list) -> int:
    """
    Find the number of "OKPO"
    :param records: data list
    :return: count sheets
    """
    list_okpo = []
    for record in records[1:]:
        if record[0] not in list_okpo:
            if record[0]:
                okpo = str(int(record[0]))
                list_okpo.append(okpo)
    return len(list_okpo)


def main():
    start = perf_counter()
    """Open xml"""
    tree = ElT.parse(config.NAME_XML_FILE)
    root = tree.getroot()

    """Read xls"""
    name_xls_file = config.NAME_XLS_FILE

    """Reading list number 10 partition 9 """
    result = read_xlrd(name_xls_file, 10)

    """Finde all fils"""
    logger.info(f"count rows in sheet {find_number_of_branches(result)}")

    """Delete all section "code='5'" in xml file"""
    section5 = root[1][4]
    for row in section5.findall('row'):
        section5.remove(row)

    """Get list with row section 5"""
    list_section5 = create_rows_section_code5(result)

    """Create all section "code='5'" in xml file"""
    for section_row in list_section5:
        sort_et = ElT.SubElement
        row = sort_et(section5, "row", attrib=section_row[0])
        sort_et(row, "col", attrib={"code": "1"}).text = str(section_row[1])
        sort_et(row, "col", attrib={"code": "2"}).text = str(section_row[2])
        sort_et(row, "col", attrib={"code": "3"}).text = str(section_row[3])
        sort_et(row, "col", attrib={"code": "4"}).text = str(section_row[4])

    """Create xml"""
    tree.write(config.CREATE_XML_NAME, encoding="WINDOWS-1251")

    """Create pritty xml"""
    ElT.indent(tree, space=" ", level=0)
    tree.write(config.CREATE_XML_NAME_2, encoding="WINDOWS-1251")

    stop = perf_counter() - start
    logger.info(f"Execution time {stop}")


if __name__ == '__main__':
    main()
