import os
import json
import win32com.client

from pywinauto import application
import time

def read_JSON(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

DASHIN_PATH  = 'C:/DASHIN'
SYSTEM_PATH = os.getcwd()

def connect(reconnect=True):

    # login_history = f'{DASHIN_PATH}/ncfsys/data//cct.dat'
    # if os.path.isfile(login_history):
    #     os.remove(f'{DASHIN_PATH}/ncfsys/data/cct.dat')
    # print('\n 자동 로그인 설정 파일 삭제 완료\n')

    login_info = read_JSON(f'{SYSTEM_PATH}/config/config.json')

    # 재연결이라면 기존 연결을 강제로 kill
    if reconnect:
        try:
            os.system('taskkill /IM ncStarter* /F /T')
            os.system('taskkill /IM CpStart* /F /T')
            os.system('taskkill /IM DibServer* /F /T')
            os.system('wmic process where "name like \'%ncStarter%\'" call terminate')
            os.system('wmic process where "name like \'%CpStart%\'" call terminate')
            os.system('wmic process where "name like \'%DibServer%\'" call terminate')
        except:
            pass

    CpCybos = win32com.client.Dispatch("CpUtil.CpCybos")

    if CpCybos.IsConnect:
        print('시스템 명령 실행 알림 - Cybos already connected...')

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

        print('시스템 명령 실행 알림 - Cybos connected...')
    return CpCybos


# 이미 연결되어있다면 재연결 x
CpCybos = connect(True)

