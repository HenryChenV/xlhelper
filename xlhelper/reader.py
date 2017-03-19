# -*- coding=utf-8 -*-
from __future__ import unicode_literals
"""
    Excel Reader
    ~~~~~~~~~~~~~
"""
import xlrd

from .base import cached_property, _missing


__all__ = ['ExcelReader']


class ExcelReader(object):

    def __init__(self, filename=None, file_contents=None, file_point=None):
        self.filename = filename
        self.file_contents = file_contents
        self.file_point = file_point

    @cached_property
    def workbook(self):
        if self.file_point:
            self.file_contents = self.file_point.read()
        return xlrd.open_workbook(filename=self.filename,
                                  file_contents=self.file_contents)

    def parse_header(self, field_descs, header):
        """解析表头

        :param field_descs: field descs, eg. (field1, field2, field3)
        :param header: header row
        :return: ((field, xl_index),
                     (<Field0 object>, 0), (<Field1 object>, 1), )
                     if field is missing, xl_index = _missing
        """
        header_descs = []
        col_index_mapping = {c.value: i for i, c in enumerate(header)}
        for field in field_descs:
            header_descs.append(
                (field, col_index_mapping.get(field.xl_name, _missing)))
        return header_descs

    def parse_sheet_data(self, field_descs, sheet_index=0):
        """解析sheet数据

        :param field_descs: 字段描述
        :sheet_index: sheet下标
        """
        records = []
        sheet = self.workbook.sheets()[sheet_index]
        nrows = sheet.nrows

        headers = self.parse_header(field_descs, sheet.row(0))
        for i_row in xrange(1, nrows):
            row = sheet.row(i_row)
            record = {}
            for field, i_col in headers:
                if i_col is not _missing:
                    record[field.key] = field(row[i_col])
                else:
                    record[field.key] = field()
            records.append(record)
        return records
