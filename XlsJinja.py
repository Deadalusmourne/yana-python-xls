#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2018/1/18 
# __author__: caoge
import re


class XlsJinja:
    """
    预设的jinja对象，一是保存xls中的变量和信息，二是customize的filter
    """
    def __init__(self):
        self.status = 0      # 循环的信息 00000000 是否循环 循环几次 剩余次数 循环宽度
    # 匹配规则
    variable_re = re.compile('\{\{(.*)\}\}')          # 变量
    variable_escape = re.compile('__\{\{(.*)\}\}__')
    control_re = re.compile('\{%(.*)%\}')             # 控制语句
    control_escape = re.compile('__\{%(.*)%\}__')
    is_formula_re = re.compile('\+|-|\*|/|%')
    forloop_formula = re.compile('^(?i)for\s+(\w+)\s+?in\s+(\w)')
    set_re = re.compile('(\w+)\s*=\s*(\w+)')
    forloop_endfor = re.compile('^endfor')
    # 参数
    tr_loop = 0       # 记录row循环开始的index
    tc_loop = 0
    render_vb = {}

    @classmethod
    def assert_text(cls, text):
        resp = {'type': '', 'data': [], 'error': ''}
        if isinstance(text, unicode):
            result_es1 = cls.control_escape.search(text)
            result_es2 = cls.variable_escape.search(text)
            if result_es1 or result_es2:
                return
            result1 = cls.control_re.search(text)
            result2 = cls.variable_re.search(text)
            if result1:
                control_jijna = result1.groups()[0] if result1.groups() else ''
                if control_jijna and isinstance(control_jijna, unicode):
                    control_jijna = control_jijna.strip()
                    if control_jijna.startswith('tr '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = cls.forloop_formula.search(control_jijna)
                        req_endfor = cls.forloop_endfor.search(control_jijna)
                        if req_for:
                            i, vb = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'tr'
                            resp['data'] = [i, vb]
                        elif req_endfor:
                            resp['type'] = 'trendfor'
                    elif control_jijna.startswith('tc '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = cls.forloop_formula.search(control_jijna)
                        req_endfor = cls.forloop_endfor.search(control_jijna)
                        if req_for:
                            i, vb = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'tc'
                            resp['data'] = [i, vb]
                        elif req_endfor:
                            resp['type'] = 'trendfor'
                    elif control_jijna.startswith('set '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = cls.set_re.search(control_jijna)
                        if req_for:
                            vb, vb_value = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'set'
                            resp['data'] = [vb, vb_value]
                    else:
                        resp['error'] = '不支持的语句'
            elif result2:
                resp['type'] = 'variable'
                resp['data'] = str(result2.groups()[0]).strip()
        return resp

    def setbit(self, offset, value):
        a = list(bin(self.status)).reverse()
        a[offset] = str(value)
        a.reverse()
        a = ''.join(a)
        self.status = int(a, 2)

    def getbit(self, offset):
        a = list(bin(self.status)).reverse()
        return a[offset]

    def render(self, render_vb):
        self.render_vb = render_vb

class MultipleIterationError(Exception):
    """
    多次迭代的控制循环语句使得Excel样式变乱
    """
    def __init__(self, test=''):
        self.message = test

    def __str__(self):
        if self.message:
            return self.message
        else:
            return 'unsupported for using loop statement more than twice'


if __name__ == '__main__':
    # raise MultipleIterationError()
    a = XlsJinja()
    print XlsJinja.__dict__

