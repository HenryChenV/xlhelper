# -*- coding=utf-8 -*-
from xlhelper import ExcelReader, fields
import pprint

field_descs = (
    fields.Int(xl_name=u'加盟商ID', key='ops_org_id', required=True,
               nullable=False),
    fields.Str(xl_name=u'加盟商名称', key='org_name', required=True,
               nullable=False),
    fields.Float(xl_name=u'金额', key='amount', required=True,
                 nullable=False, as_decimal=True),
    fields.Str(xl_name=u'备注', key='remark', required=True,
               nullable=True, default='')
)

reader = ExcelReader(
    filename='/Users/henry/Desktop/线下申款测试数据-加盟商.xlsx')

try:
    rv = reader.parse_sheet_data(field_descs)
    pprint.pprint(rv)
except Exception as e:
    import traceback
    traceback.print_exc()
    import ipdb
    ipdb.set_trace()
