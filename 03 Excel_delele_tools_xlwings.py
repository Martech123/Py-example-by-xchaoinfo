#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Date    : 2017-03-16 14:14:03
# @Author  : xchaoinfo (xchaoinfo)
# @github  : https://github.com/xchaoinfo

import xlwings as xw
import numpy as np
import pandas as pd

#fn = "r.xlsx"
class CapacitorExcelEdit(object):
    def __init__(self, fn):
        super(CapacitorExcelEdit, self).__init__()
        self.ExistSet = set()
        self.ToDelList = list()
        self.fn = fn

    def val_change(self, value):
        print(type(value))
        if ((value-int(value)) != 0):
            if value > 1:
                print("it's a float")
                new_value = str(value).replace('.', 'P')
                return  new_value
            else:
                print("it's a float")
                new_value = str(value) + 'P'
                return  new_value
        else:
            if value < 1000:
                new_value = str(int(value)) + 'P'
                return new_value
            elif 1000 <= value < 1000000:
                value = int(value)
                if ((value % 1000) == 0):
                    value = value / 1000
                    return (str(value) + 'N')
                elif ((value % 100) == 0):
                    value = value / 100
                    value_list = list(str(value))
                    value_list.insert(-1, 'N')
                    value_str = "".join(value_list)
                    return value_str
                elif ((value % 10) == 0):
                    value = value / 10
                    value_list = list(str(value))
                    value_list.insert(-2, 'N')
                    value_str = "".join(value_list)
                    return value_str
                else:
                    value_list = list(str(value))
                    value_list.append('P')
                    value_str = "".join(value_list)
                    return value_str
            else:
                value = int(value)
                if ((value % 1000000) == 0):
                    value = value / 1000000
                    return (str(value) + 'U')
                elif ((value % 100000) == 0):
                    value = value / 100000
                    value_list = list(str(value))
                    value_list.insert(-1, 'U')
                    value_str = "".join(value_list)
                    return value_str
                elif ((value % 10000) == 0):
                    value = value / 10000
                    value_list = list(str(value))
                    value_list.insert(-2, 'U')
                    value_str = "".join(value_list)
                    return value_str
                else:
                    value_list = list(str(value))
                    value_list.append('P')
                    value_str = "".join(value_list)
                    return value_str

    def value_tihuan(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open('c_sorted.xlsx')
        sheet = wb.sheets.active

        r_value = sheet.range("F2:F3904")
        print(r_value)
        i = 0
        for R in r_value:
            print(R.value)
            if R.value is None:
                sheet.autofit()
                wb.save('c_tihuan.xlsx')
                app.quit()
                return
            else:
                R.value = self.val_change(R.value)
                i = i + 1
                print('i:%d'%i)
        sheet.autofit()
        wb.save('c_tihuan.xlsx')
        app.quit()

    def rule(self, value):
        pass

    def emptyrow_delete(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(self.fn)
        sheet = wb.sheets.active

        r_value = sheet.range("G1:G3904")
        print(r_value)
        i = 0
        for R in r_value:
            if R.value is None:
                self.ToDelList.append(R.address)
                i = i + 1
                print('i:%d'%i)
        print(self.ToDelList)

        while self.ToDelList:
            td = self.ToDelList.pop()
            sheet.range(td).api.EntireRow.Delete()
            print(self.ToDelList)
        sheet.autofit()
        wb.save()
        app.quit()

    def value_sort(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open('c.xlsx')
        sheet = wb.sheets.active
        r_dict = {}

        r_dict['r_value'] = sheet.range("G1:G6645").value
        r_dict['r_pn'] = sheet.range("A1:A6645").value
        r_dict['r_discription'] = sheet.range("B1:B6645").value
        r_dict['r_package'] = sheet.range("C1:C6645").value
        r_dict['r_offset'] = sheet.range("D1:D6645").value
        r_dict['r_type'] = sheet.range("E1:E6645").value
        r_dict['r_power'] = sheet.range("F1:F6645").value
        #print(r_dict)
        #r_dict = {
        #            'r_value': r_value,
        #            'r_pn': r_pn,
        #            'r_discription': r_discription,
        #            'r_package': r_package,
        #            'r_offset': r_offset,
        #            'r_type': r_type,
        #            't_power': r_power,
        #            'r_dict': r_dict
        #        }
        #r_dict = {
        #        'xingming':['zhangfei', 'guanyu', 'zhaoyun', 'caocao', 'sunquan', 'liubei'],
        #        'r_value':[27, 32, 32, 39, 20, 39],
        #        'shengao':[183, 187, 183, 175, 183, 183],
        #        'wuli':[95, 97, 97, 77, 57, 82]
        #    }
        df = pd.DataFrame(r_dict)
        df = df.sort_values('r_power', ascending=True)
        print(df)
        df.to_excel('c_sorted.xlsx')
        #wb.sheets[0].range['A1'].value = df
        sheet.autofit()
        wb.save()
        app.quit()


class ResistorExcelEdit(object):
    def __init__(self, fn):
        super(ResistorExcelEdit, self).__init__()
        self.ExistSet = set()
        self.ToDelList = list()
        self.fn = fn

    def val_change(self, value):
        print(type(value))
        if ((value-int(value)) != 0):
            print("it's a float")
            new_value = str(value).replace('.', 'R')
            return  new_value
        else:
            if value < 1000:
                new_value = str(int(value)) + 'R'
                return new_value
            elif 1000 <= value < 1000000:
                value = int(value)
                if ((value % 1000) == 0):
                    value = value / 1000
                    return (str(value) + 'K')
                elif ((value % 100) == 0):
                    value = value / 100
                    value_list = list(str(value))
                    value_list.insert(-1, 'K')
                    value_str = "".join(value_list)
                    return value_str
                elif ((value % 10) == 0):
                    value = value / 10
                    value_list = list(str(value))
                    value_list.insert(-2, 'K')
                    value_str = "".join(value_list)
                    return value_str
                else:
                    value_list = list(str(value))
                    value_list.append('R')
                    value_str = "".join(value_list)
                    return value_str
            else:
                value = int(value)
                if ((value % 1000000) == 0):
                    value = value / 1000000
                    return (str(value) + 'M')
                elif ((value % 100000) == 0):
                    value = value / 100000
                    value_list = list(str(value))
                    value_list.insert(-1, 'M')
                    value_str = "".join(value_list)
                    return value_str
                elif ((value % 10000) == 0):
                    value = value / 10000
                    value_list = list(str(value))
                    value_list.insert(-2, 'M')
                    value_str = "".join(value_list)
                    return value_str
                else:
                    value_list = list(str(value))
                    value_list.append('R')
                    value_str = "".join(value_list)
                    return value_str

    def value_tihuan(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open('r_sorted.xlsx')
        sheet = wb.sheets.active

        r_value = sheet.range("H2:H6646")
        print(r_value)
        i = 0
        for R in r_value:
            print(R.value)
            if R.value is None:
                sheet.autofit()
                wb.save('r_tihuan.xlsx')
                app.quit()
                return
            else:
                R.value = self.val_change(R.value)
                i = i + 1
                print('i:%d'%i)
        sheet.autofit()
        wb.save('r_tihuan.xlsx')
        app.quit()

    def rule(self, value):
        pass

    def emptyrow_delete(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(self.fn)
        sheet = wb.sheets.active

        r_value = sheet.range("G1:G7207")
        print(r_value)
        i = 0
        for R in r_value:
            if R.value is None:
                self.ToDelList.append(R.address)
                i = i + 1
                print('i:%d'%i)
        print(self.ToDelList)

        while self.ToDelList:
            td = self.ToDelList.pop()
            sheet.range(td).api.EntireRow.Delete()
            print(self.ToDelList)
        sheet.autofit()
        wb.save()
        app.quit()

    def value_sort(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open('r.xlsx')
        sheet = wb.sheets.active
        r_dict = {}

        r_dict['r_value'] = sheet.range("G1:G6645").value
        r_dict['r_pn'] = sheet.range("A1:A6645").value
        r_dict['r_discription'] = sheet.range("B1:B6645").value
        r_dict['r_package'] = sheet.range("C1:C6645").value
        r_dict['r_offset'] = sheet.range("D1:D6645").value
        r_dict['r_type'] = sheet.range("E1:E6645").value
        r_dict['r_power'] = sheet.range("F1:F6645").value
        #print(r_dict)
        #r_dict = {
        #            'r_value': r_value,
        #            'r_pn': r_pn,
        #            'r_discription': r_discription,
        #            'r_package': r_package,
        #            'r_offset': r_offset,
        #            'r_type': r_type,
        #            't_power': r_power,
        #            'r_dict': r_dict
        #        }
        #r_dict = {
        #        'xingming':['zhangfei', 'guanyu', 'zhaoyun', 'caocao', 'sunquan', 'liubei'],
        #        'r_value':[27, 32, 32, 39, 20, 39],
        #        'shengao':[183, 187, 183, 175, 183, 183],
        #        'wuli':[95, 97, 97, 77, 57, 82]
        #    }
        df = pd.DataFrame(r_dict)
        df = df.sort_values('r_value', ascending=True)
        print(df)
        df.to_excel('r_sorted.xlsx')
        #wb.sheets[0].range['A1'].value = df
        sheet.autofit()
        wb.save()
        app.quit()

if __name__ == '__main__':
    #d = ResistorExcelEdit(fn)
    #d.emptyrow_delete()
    #d.value_sort()
    #d.value_tihuan()

    d = CapacitorExcelEdit('c.xlsx')
    #d.emptyrow_delete()
    #d.value_sort()
    d.value_tihuan()
