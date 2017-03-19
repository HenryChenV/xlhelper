# -*- coding=utf-8 -*-
"""
    Excel Field Desc
    ~~~~~~~~~~~~~~~~
"""
import sys
from xlrd import biffh
from xlrd.sheet import Cell
from decimal import Decimal

from .exception import BaseException
from .base import _missing


__all__ = ['XL_CELL_TYPE', 'is_missing', 'is_empty', 'Int', 'Float', 'Str']


class XL_CELL_TYPE(object):
    # inherit from xlrd
    EMPTY = biffh.XL_CELL_EMPTY
    TEXT = biffh.XL_CELL_TEXT
    NUMBER = biffh.XL_CELL_NUMBER
    DATE = biffh.XL_CELL_DATE
    BOOLEAN = biffh.XL_CELL_BOOLEAN
    ERROR = biffh.XL_CELL_ERROR
    BLANK = biffh.XL_CELL_BLANK


class InvalidField(BaseException):

    def __init__(self, message=''):
        super(InvalidField, self).__init__('invalid_field', message=message)


def is_missing(val):
    return val is _missing


def is_empty(cell):
    return cell.ctype == XL_CELL_TYPE.EMPTY


class Field(object):
    """字段

    :param xl_name: 字段名称
    :param xl_ctype: 字段类型
    :param key: 字段输出时的key
    :param as_str: str(output_value)
    :param required: is required?
    :param nullable: is nullable?
    :param missing: if field is missing, user missing insteaded
    :param default: if field value is empty, user default instead
    """

    xl_ctype = None
    cell_cls = Cell

    def __init__(self, xl_name, key=_missing, as_str=False,
                 required=False, default=_missing, missing=_missing,
                 nullable=True):
        self.xl_name = xl_name
        self.key = key or xl_name
        self.as_str = as_str
        self.required = required
        self.nullable = nullable
        self.default = default if is_missing(default) else self._2cell(default)
        self.missing = missing if is_missing(missing) else self._2cell(missing)

    def _2cell(self, val):
        """ 把val封装成cell
        """
        if val in (u'', '', None):
            return self.cell_cls(ctype=XL_CELL_TYPE.EMPTY, value='')
        return self.cell_cls(ctype=self.xl_ctype, value=val)

    def is_required(self):
        return self.required

    def is_nullable(self):
        return self.nullable

    def format_(self, val):
        if self.as_str:
            return str(val)
        return val

    def validate_required(self, cell):
        if self.required:
            if is_missing(cell):
                raise InvalidField('field %s is required, but %s!' % (
                    self.__class__.__name__.lower(), cell
                ))
        if is_missing(cell) and not is_missing(self.missing):
            cell = self.missing
        return cell

    def validate_nullable(self, cell):
        if not self.nullable:
            if is_empty(cell):
                raise InvalidField(message='field %s is not nullable, but %s!' \
                                   % (self.__class__.__name__.lower(), cell))
        if cell.ctype == XL_CELL_TYPE.EMPTY and not is_missing(self.default):
            cell = self.default
        return cell

    def validate(self, cell):
        """验证cell是否是字段类型
        """
        # required?
        cell = self.validate_required(cell)
        # nullable?
        cell = self.validate_nullable(cell)
        if is_empty(cell):
            return cell.value
        # type?
        if cell.ctype != self.xl_ctype:
            raise InvalidField(message='invalid field %s: %s!' % (
                self.__class__.__name__.lower(), cell))
        return self.format_(cell.value)

    def __call__(self, cell=_missing):
        """返回cell的值
        """
        return self.validate(cell)

# xlrd fields: Emapty, Text, Number, Date, Boolean, Error, Blank
this_module = sys.modules[__name__]
xlrd_fields = ('Emapty', 'Text', 'Number', 'Date', 'Boolean', 'Error', 'Blank')
for class_name in xlrd_fields:
    if not getattr(this_module, class_name, None):
        setattr(this_module, class_name,
                type(class_name, (Field, ), \
                     {'xl_ctype': getattr(
                         XL_CELL_TYPE, class_name.upper(), None)}))
__all__.extend(xlrd_fields)


class Int(Number):

    def format_(self, val):
        val = int(val)
        return super(Int, self).format_(val)


class Float(Number):
    """Float

    :param as_decimal: Decimal(output_value)
    """

    def __init__(self, xl_name, key=_missing, as_str=False, nullable=True,
                 required=False,  default=_missing, as_decimal=True,
                 missing=_missing):
        as_str = True if not as_str and not as_decimal else as_str
        super(Float, self).__init__(xl_name, key=key, as_str=as_str,
                                    nullable=nullable, required=required,
                                    default=default, missing=missing)
        self.as_decimal = as_decimal

    def format_(self, val):
        if self.as_decimal:
            return Decimal(str(val))
        return super(Float, self).format_(val)


class Str(Field):
    """Str

    一种特殊的类型，所有的类型都可以看做是str，所以对cell的指不做校验，直接返回
    """

    def validate(self, cell):
        # no validation
        return cell.value
