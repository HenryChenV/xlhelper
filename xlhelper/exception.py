# -*- coding=utf-8 -*-


class BaseException(Exception):

    def __init__(self, code, message=''):
        self.code = code
        self.message = message

    def __str__(self):
        return ('<{module_name}.{class_name}(code={code}, '
                'message={message})>').format(
                    module_name=self.__class__.__module__,
                    class_name=self.__class__.__name__,
                    code=self.code,
                    message=self.message
                )
    __repr__ = __str__
