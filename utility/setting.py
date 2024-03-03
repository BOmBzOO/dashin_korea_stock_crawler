import os
import json
import sqlite3
from PyQt5.QtGui import QFont, QColor


K_OPENAPI_PATH = 'C:/OpenAPI'
E_OPENAPI_PATH = 'C:/eBEST/xingAPI'

# SYSTEM_PATH = 'E:/xing'
SYSTEM_PATH = os.getcwd()
ICON_PATH = f'{SYSTEM_PATH}/utility/icon'
DB_SETTING = f'{SYSTEM_PATH}/database/setting.db'

DB_STOCK_TICK = f'{SYSTEM_PATH}/database/stock_tick.db'
DB_STOCK_BROKER = f'{SYSTEM_PATH}/database/stock_broker.db'
DB_STOCK_VI = f'{SYSTEM_PATH}/database/stock_vi.db'

# SYSTEM_PATH = 'C:/data'
DB_STOCK_HOGA = f'{SYSTEM_PATH}/database/stock_hoga.db'


DB_FUTURE_TICK = f'{SYSTEM_PATH}/database/future_tick.db'
DB_FUTURE_HOGA = f'{SYSTEM_PATH}/database/future_hoga.db'

DICT_SET = {
    '저장주기': 60,
}

qfont12 = QFont()
qfont12.setFamily('나눔고딕')
qfont12.setPixelSize(12)

qfont14 = QFont()
qfont14.setFamily('나눔고딕')
qfont14.setPixelSize(14)

color_fg_bt = QColor(230, 230, 235)
color_fg_bc = QColor(190, 190, 195)
color_fg_dk = QColor(150, 150, 155)
color_fg_bk = QColor(110, 110, 115)
color_fg_hl = QColor(110, 110, 255)

color_bg_bt = QColor(50, 50, 55)
color_bg_bc = QColor(40, 40, 45)
color_bg_dk = QColor(30, 30, 35)
color_bg_bk = QColor(20, 20, 25)

color_bf_bt = QColor(110, 110, 115)
color_bf_dk = QColor(70, 70, 75)

color_cs_hr = QColor(230, 230, 0)

style_fc_bt = 'color: rgb(230, 230, 235);'
style_fc_dk = 'color: rgb(150, 150, 155);'
style_bc_st = 'background-color: rgb(70, 70, 75);'
style_bc_bt = 'background-color: rgb(50, 50, 55);'
style_bc_dk = 'background-color: rgb(30, 30, 35);'
style_bc_by = 'background-color: rgb(100, 70, 70);'
style_bc_sl = 'background-color: rgb(70, 70, 100);'
style_pgbar = 'QProgressBar {background-color: #28282d;} QProgressBar::chunk {background-color: #5a5a5f;}'

ui_num = {'설정텍스트': 0, 'S단순텍스트': 1, 'S로그텍스트': 2, 'S종목명딕셔너리': 3,
          'C단순텍스트': 4, 'C로그텍스트': 5, 'S백테스트': 6, 'C백테스트': 7,
          'S실현손익': 11, 'S거래목록': 12, 'S잔고평가': 13, 'S잔고목록': 14, 'S체결목록': 15,
          'S당일합계': 16, 'S당일상세': 17, 'S누적합계': 18, 'S누적상세': 19, 'S관심종목': 20,
          'C실현손익': 21, 'C거래목록': 22, 'C잔고평가': 23, 'C잔고목록': 24, 'C체결목록': 25,
          'C당일합계': 26, 'C당일상세': 27, 'C누적합계': 28, 'C누적상세': 29, 'C관심종목': 30,
          '차트': 40, '실시간차트': 41,
          'S호가종목': 42, 'S호가체결': 43, 'S호가잔량': 44,
          'C호가종목': 45, 'C호가체결': 46, 'C호가잔량': 47}

columns_tt = ['거래횟수', '총매수금액', '총매도금액', '총수익금액', '총손실금액', '수익률', '수익금합계']
columns_td = ['종목명', '매수금액', '매도금액', '주문수량', '수익률', '수익금', '체결시간']
columns_tj = ['추정예탁자산', '추정예수금', '보유종목수', '수익률', '총평가손익', '총매입금액', '총평가금액']
columns_jg = ['종목명', '매입가', '현재가', '수익률', '평가손익', '매입금액', '평가금액', '보유수량']
columns_cj = ['종목명', '주문구분', '주문수량', '미체결수량', '주문가격', '체결가', '체결시간']
columns_gj = ['등락율', '고저평균대비등락율', '초당거래대금', '초당거래대금평균', '당일거래대금',
              '체결강도', '체결강도평균', '최고체결강도', '현재가']
columns_gj_ = ['종목명', 'per', 'hlml_per', 's_money', 'sm_avg', 'd_money', 'ch', 'ch_avg', 'ch_high']

columns_dt = ['거래일자', '누적매수금액', '누적매도금액', '누적수익금액', '누적손실금액', '수익률', '누적수익금']
columns_dd = ['체결시간', '종목명', '매수금액', '매도금액', '주문수량', '수익률', '수익금']
columns_nt = ['기간', '누적매수금액', '누적매도금액', '누적수익금액', '누적손실금액', '수익률', '누적수익금']
columns_nd = ['일자', '총매수금액', '총매도금액', '총수익금액', '총손실금액', '수익률', '수익금합계']

columns_sm = ['증권사', '주식리시버', '주식콜렉터', '주식트레이더', '거래소', '코인리시버', '코인콜렉터', '코인트레이더', '주식순위시간',
              '주식순위선정', '코인순위시간', '코인순위선정', '주식실시간저장', '주식전체종목저장', '주식저장주기', '코인저장주기']
columns_sk = ['아이디1', '비밀번호1', '인증서비밀번호1', '계좌비밀번호1', '아이디2', '비밀번호2', '인증서비밀번호2', '계좌비밀번호2']
columns_su = ['Access_key', 'Secret_key']
columns_st = ['str_bot', 'int_id']

columns_hj = ['종목명', '현재가', '등락율', 'UVI', '시가', '고가', '저가']
columns_hc = ['체결수량', '체결강도']
columns_hg = ['잔량', '호가']
