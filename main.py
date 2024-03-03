import sys
import json
import psutil
import time
import logging
import pyqtgraph as pg
import pandas as pd

import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
import win32com.client
from PyQt5.QtCore import Qt
from PyQt5 import QtCore, QtGui, QtWidgets
from multiprocessing import Process, get_context
from utility.setui import *
from utility.setting import *
from pywinauto import application
from utility.utility import now, strf_time, read_JSON, timedelta_sec, float2str1p6
from creon_datareader_cli import CreonDatareaderCLI

DASHIN_PATH  = 'C:/DASHIN'
SYSTEM_PATH = os.getcwd()

class Writer(QtCore.QThread):
    data1 = QtCore.pyqtSignal(list)

    def __init__(self):
        super().__init__()

    def run(self):
        while True:
            data = windowQ.get()
            if data[0] <= 10:
                self.data1.emit(data)

class Window(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.log1 = logging.getLogger('Stock')
        self.log1.setLevel(logging.INFO)
        filehandler = logging.FileHandler(filename=f"{SYSTEM_PATH}/log/S{strf_time('%Y%m%d')}.txt", encoding='utf-8')
        self.log1.addHandler(filehandler)
        
        SetUI(self)

        self.cybos_started = False
        self.CpCybos = None
        self.time_queue_print = now()
        self.Collector = CreonDatareaderCLI(windowQ, ui_num)

        self.qtimer1 = QtCore.QTimer()
        self.qtimer1.setInterval(1000)
        self.qtimer1.timeout.connect(self.collector)
        self.qtimer1.start()

        self.writer = Writer()
        self.writer.data1.connect(self.UpdateTexedit)
        self.writer.start()
    
    def __del__(self):
        self.Collector.stop()
        self.Collector.con.close()
        if self.qtimer1.isActive():
            self.qtimer1.stop()
        if self.writer.isRunning():
            self.writer.terminate()


    def collector(self):
        if not (830000 < int(strf_time('%H%M%S')) > 170000) and not self.cybos_started:
            self.cybos_started = True

            db_list = {
                'Day':['./db/stock_price(day).db', 'day', True],
                'Tick':['./db/stock_price(tick).db', 'tick', False],
                '1Min':['./db/stock_price(1min).db', '1min', False],
                '5Min':['./db/stock_price(5min).db', '5min', False]
            }
            for key in db_list.keys():
                massage = f'시스템 명령 실행 알림 - {key} 데이터 수집 시작'
                windowQ.put([ui_num['S로그텍스트'], massage])
                print(massage)

                self.Collector.update_price_db(db_list[key][0], tick_unit=db_list[key][1], ohlcv_only=db_list[key][2])
                massage = f'시스템 명령 실행 알림 - {key} 데이터 수집 완료'
                windowQ.put([ui_num['S로그텍스트'], massage])
                print(massage)

            self.Collector.stop()
            self.Collector.con.close()
            if self.qtimer1.isActive():
                self.qtimer1.stop()
            if self.writer.isRunning():
                self.writer.terminate()
                
        if 830000 < int(strf_time('%H%M%S')) > 170000:
            self.cybos_started = False

    def UpdateTexedit(self, data):
        text = f'[{now()}]  {data[1]}'
        if data[0] == ui_num['S로그텍스트']:
            self.log_lu_textEdit.append(text)
            self.log1.info(text)
        elif data[0] == ui_num['S단순텍스트']:
            self.log_ld_textEdit.append(text)

    # noinspection PyArgumentList
    def closeEvent(self, a):
        buttonReply = QtWidgets.QMessageBox.question(
            self, "프로그램 종료", "프로그램을 종료합니다.",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No
        )
        if buttonReply == QtWidgets.QMessageBox.Yes:
            if self.qtimer1.isActive():
                self.qtimer1.stop()
            if self.writer.isRunning():
                self.writer.terminate()
            self.Collector.stop()
            a.accept()
        else:
            a.ignore()

def read_JSON(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

if __name__ == "__main__":

    ctx = get_context("spawn")
    windowQ = ctx.Queue()

    login_info = read_JSON(f'{SYSTEM_PATH}/config/config.json')
    CpCybos = win32com.client.Dispatch("CpUtil.CpCybos")

    if CpCybos.IsConnect:
        massage = f'시스템 명령 실행 알림 - Cybos already connected...'
        windowQ.put([ui_num['S로그텍스트'], massage])
        print(massage)

    else:
        app = application.Application()
        app.start(
            'C:/DAISHIN/STARTER/ncStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                id=login_info['id'], pwd=login_info['pwd'], pwdcert=login_info['pwdcert'])
        )
        # 연결 될때까지 무한루프
        while True:
            if CpCybos.IsConnect:
                break
            time.sleep(1)

        massage = f'시스템 명령 실행 알림 - Cybos connected...'
        windowQ.put([ui_num['S로그텍스트'], massage])
        print(massage)
        # print('시스템 명령 실행 알림 - Cybos connected...')

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle(ProxyStyle())
    app.setStyle('fusion')
    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, color_bg_bc)
    palette.setColor(QtGui.QPalette.Background, color_bg_bc)
    palette.setColor(QtGui.QPalette.WindowText, color_fg_bc)
    palette.setColor(QtGui.QPalette.Base, color_bg_bc)
    palette.setColor(QtGui.QPalette.AlternateBase, color_bg_dk)
    palette.setColor(QtGui.QPalette.Text, color_fg_bc)
    palette.setColor(QtGui.QPalette.Button, color_bg_bc)
    palette.setColor(QtGui.QPalette.ButtonText, color_fg_bc)
    palette.setColor(QtGui.QPalette.Link, color_fg_bk)
    palette.setColor(QtGui.QPalette.Highlight, color_fg_hl)
    palette.setColor(QtGui.QPalette.HighlightedText, color_bg_bk)
    app.setPalette(palette)
    window = Window()
    window.show()
    app.exec_()





