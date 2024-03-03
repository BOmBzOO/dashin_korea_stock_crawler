import pyqtgraph
from PyQt5.QtCore import Qt
from PyQt5 import QtCore, QtGui, QtWidgets
# from utility import syntax
from utility.static import CustomViewBox
from utility.setting import qfont12, qfont14, style_bc_st, style_bc_bt, style_bc_dk, style_fc_bt, style_pgbar, \
    columns_tt, columns_td, columns_tj, columns_jg, columns_gj_, columns_cj, columns_dt, columns_dd, columns_nt, \
    columns_nd, ICON_PATH, style_bc_by, style_bc_sl, columns_hj, columns_hc, columns_hg, style_fc_dk

class TabBar(QtWidgets.QTabBar):
    def tabSizeHint(self, index):
        s = QtWidgets.QTabBar.tabSizeHint(self, index)
        s.setWidth(40)
        s.setHeight(40)
        s.transpose()
        return s

    def paintEvent(self, event):
        painter = QtWidgets.QStylePainter(self)
        opt = QtWidgets.QStyleOptionTab()

        for i in range(self.count()):
            self.initStyleOption(opt, i)
            painter.drawControl(QtWidgets.QStyle.CE_TabBarTabShape, opt)
            painter.save()

            s = opt.rect.size()
            s.transpose()
            r = QtCore.QRect(QtCore.QPoint(), s)
            r.moveCenter(opt.rect.center())
            opt.rect = r

            c = self.tabRect(i).center()
            painter.translate(c)
            painter.rotate(90)
            painter.translate(-c)
            painter.drawControl(QtWidgets.QStyle.CE_TabBarTabLabel, opt)
            painter.restore()


class TabWidget(QtWidgets.QTabWidget):
    def __init__(self, *args, **kwargs):
        QtWidgets.QTabWidget.__init__(self, *args, **kwargs)
        self.setTabBar(TabBar(self))
        self.setTabPosition(QtWidgets.QTabWidget.West)

class ProxyStyle(QtWidgets.QProxyStyle):
    def drawControl(self, element, opt, painter, widget=None):
        if element == QtWidgets.QStyle.CE_TabBarTabLabel:
            ic = self.pixelMetric(QtWidgets.QStyle.PM_TabBarIconSize)
            r = QtCore.QRect(opt.rect)
            w = 0 if opt.icon.isNull() else opt.rect.width() + ic
            r.setHeight(opt.fontMetrics.width(opt.text) + w)
            r.moveBottom(opt.rect.bottom())
            opt.rect = r
        QtWidgets.QProxyStyle.drawControl(self, element, opt, painter, widget)


def SetUI(self):

    # def setPushbutton(name, box=None, click=None, cmd=None, icon=None, tip=None, color=0):
    #     if box is not None:
    #         pushbutton = QtWidgets.QPushButton(name, box)
    #     else:
    #         pushbutton = QtWidgets.QPushButton(name, self)
    #     if color == 0:
    #         pushbutton.setStyleSheet(style_bc_bt)
    #     elif color == 1:
    #         pushbutton.setStyleSheet(style_bc_st)
    #     elif color == 2:
    #         pushbutton.setStyleSheet(style_bc_by)
    #     elif color == 3:
    #         pushbutton.setStyleSheet(style_bc_sl)
    #     pushbutton.setFont(qfont12)
    #     if click is not None:
    #         if cmd is not None:
    #             pushbutton.clicked.connect(lambda: click(cmd))
    #         else:
    #             pushbutton.clicked.connect(click)
    #     if icon is not None:
    #         pushbutton.setIcon(icon)
    #     if tip is not None:
    #         pushbutton.setToolTip(tip)
    #     return pushbutton

    def setTextEdit(tab):
        textedit = QtWidgets.QTextEdit(tab)
        textedit.setReadOnly(True)
        textedit.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        textedit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        textedit.setStyleSheet(style_bc_dk)
        return textedit

    icon_main = QtGui.QIcon(f'{ICON_PATH}/python.png')
    icon_stock = QtGui.QIcon(f'{ICON_PATH}/stock.png')
    icon_coin = QtGui.QIcon(f'{ICON_PATH}/coin.png')
    icon_set = QtGui.QIcon(f'{ICON_PATH}/set.png')
    icon_log = QtGui.QIcon(f'{ICON_PATH}/log.png')
    icon_total = QtGui.QIcon(f'{ICON_PATH}/total.png')
    icon_start = QtGui.QIcon(f'{ICON_PATH}/start.png')
    icon_zoom = QtGui.QIcon(f'{ICON_PATH}/zoom.png')
    icon_dbdel = QtGui.QIcon(f'{ICON_PATH}/dbdel.png')
    icon_accdel = QtGui.QIcon(f'{ICON_PATH}/accdel.png')
    icon_stocks = QtGui.QIcon(f'{ICON_PATH}/stocks.png')
    icon_coins = QtGui.QIcon(f'{ICON_PATH}/coins.png')

    self.setFont(qfont14)
    self.setWindowTitle('대신증권 주식데이터 수집기')

    self.main_tabWidget = TabWidget(self)
    self.log_tab = QtWidgets.QWidget()

    self.main_tabWidget.addTab(self.log_tab, '')
    self.main_tabWidget.setTabIcon(0, icon_log)
    self.main_tabWidget.setTabToolTip(0, '  로그')
    
    self.setFixedSize(800, 763)
    self.geometry().center()
    self.main_tabWidget.setGeometry(5, 5, 790, 753)

    self.log_lu_textEdit = setTextEdit(self.log_tab)
    self.log_ld_textEdit = setTextEdit(self.log_tab)
  
    self.log_lu_textEdit.setGeometry(5, 5, 768, 367)
    self.log_ld_textEdit.setGeometry(5, 377, 768, 367)
   
    

    