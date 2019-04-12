#! /usr/local/bin/python3

import openpyxl
import datetime

"""
nippo.py
日報ファイルの準備スクリプト
Copyright (C) 2019 J.Kawahara
2019.4.11 J.Kawahara 新規作成
"""


def prepare_nippo(filename):
    """
    日報ファイルの準備を行う

    Parameters
    ----------
    filename : str
        処理対象のファイル名
    """
    nowdt = datetime.date.today()
    wb = openpyxl.load_workbook(filename)
    if wb is None:
        return
    # ワークシートをコピーする
    todays_sheet = wb.copy_worksheet(wb['雛形'])
    todays_sheet.title = nowdt.strftime('%Y-%m-%d')
    wb._sheets.sort(key=lambda ws: ws.title, reverse=True)

    # 日付を入力する
    serial_value = nowdt - datetime.date(1900, 1, 1)
    todays_sheet['H1'].value = serial_value.days

    # 保存する
    wb.save(filename)


if __name__ == '__main__':
    prepare_nippo('test.xlsx')
