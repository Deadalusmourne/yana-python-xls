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
        self.status = '00000000'      # 循环的信息 00000000 是否循环 循环几次 剩余次数 循环宽度
    # 匹配规则
    variable_re = re.compile('\{\{(.*)\}\}')          # 变量
    variable_escape = re.compile('__\{\{(.*)\}\}__')
    control_re = re.compile('\{%(.*)%\}')             # 控制语句
    control_escape = re.compile('__\{%(.*)%\}__')
    is_formula_re = re.compile('\+|-|\*|/|%')
    forloop_formula = re.compile('^(?i)for\s+(\w+)\s+?in\s+(\w+)')   # legacy for i in vb.data ?
    set_re = re.compile('(\w+)\s*=\s*(\w+)')
    forloop_endfor = re.compile('^endfor')
    # 参数
    tr_loop = 0       # 记录row循环开始的index
    tc_loop = 0
    render_vb = {}    # 默认render的参数
    tr_loop_temp_vb = []
    tc_loop_temp_vb = []

    def assert_text(self, text):
        resp = {'type': '', 'data': [], 'error': ''}
        if isinstance(text, unicode):
            result_es1 = self.control_escape.search(text)
            result_es2 = self.variable_escape.search(text)
            if result_es1 or result_es2:
                return
            result1 = self.control_re.search(text)
            result2 = self.variable_re.search(text)
            if result1:
                control_jijna = result1.groups()[0] if result1.groups() else ''
                if control_jijna and isinstance(control_jijna, unicode):
                    control_jijna = control_jijna.strip()
                    if control_jijna.startswith('tr '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = self.forloop_formula.search(control_jijna)
                        req_endfor = self.forloop_endfor.search(control_jijna)
                        if req_for:
                            i, vb = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'tr'
                            resp['data'] = [i, vb]
                        elif req_endfor:
                            resp['type'] = 'trendfor'
                    elif control_jijna.startswith('tc '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = self.forloop_formula.search(control_jijna)
                        req_endfor = self.forloop_endfor.search(control_jijna)
                        if req_for:
                            i, vb = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'tc'
                            resp['data'] = [i, vb]
                        elif req_endfor:
                            resp['type'] = 'trendfor'
                    elif control_jijna.startswith('set '):
                        control_jijna = control_jijna[2:].strip()
                        req_for = self.set_re.search(control_jijna)
                        if req_for:
                            vb, vb_value = req_for.groups() if req_for.groups() else ['', '']
                            resp['type'] = 'set'
                            resp['data'] = [vb, vb_value]
                    else:
                        resp['error'] = '不支持的语句'
            elif result2:
                resp['type'] = 'variable'
                vb_string = str(result2.groups()[0]).strip()
                print 'vb_string', vb_string
                if self.getbit(0):
                    start_str = ''.join([self.tr_loop_temp_vb[0], '.'])
                    if vb_string.startswith(start_str):
                        vb_string = vb_string[len(self.tr_loop_temp_vb[0])+1:]
                    resp['tr_or_tc'] = 0    # 判断是那个循环的变量
                    resp['loop_vb'] = self.tr_loop_temp_vb[1]
                if self.getbit(1):
                    start_str = ''.join([self.tc_loop_temp_vb[0], '.'])
                    if vb_string.startswith(start_str):
                        vb_string = vb_string[len(self.tr_loop_temp_vb[0])+1:]
                    resp['tr_or_tc'] = 1  # 判断是那个循环的变量
                    resp['loop_vb'] = self.tr_loop_temp_vb[1]
                resp['data'] = vb_string

        return resp

    def setbit(self, offset, value):
        a = list(self.status)
        a[offset] = str(value)
        a = ''.join(a)
        self.status = a

    def getbit(self, offset):
        a = list(self.status)
        return int(a[offset])

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

    # error处理
    # 1 控制语句写成了{{}} 导致的报错
    # 2 当context不存在某个循环体的时候

    # 尚未处理
    # 1 重置status
    # 2 循环体内部处理