import os
import json
import datetime
from threading import Thread
from multiprocessing import Process, Queue, freeze_support, get_context

def read_JSON(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def thread_decorator(func):
    def wrapper(*args):
        Thread(target=func, args=args, daemon=True).start()
    return wrapper


def now():
    return datetime.datetime.now()


def timedelta_sec(second, std_time=None):
    if std_time is None:
        next_time = now() + datetime.timedelta(seconds=second)
    else:
        next_time = std_time + datetime.timedelta(seconds=second)
    return next_time


def timedelta_hour(hour, std_time=None):
    if std_time is None:
        next_time = now() + datetime.timedelta(hours=hour)
    else:
        next_time = std_time + datetime.timedelta(hours=hour)
    return next_time


def timedelta_day(day, std_time=None):
    if std_time is None:
        next_time = now() + datetime.timedelta(days=day)
    else:
        next_time = std_time + datetime.timedelta(days=day)
    return next_time


def strp_time(timetype, str_time):
    return datetime.datetime.strptime(str_time, timetype)


def strf_time(timetype, std_time=None):
    if std_time is None:
        str_time = now().strftime(timetype)
    else:
        str_time = std_time.strftime(timetype)
    return str_time


def changeFormat(text, dotdowndel=False, dotdown8=False):
    text = str(text)
    try:
        format_data = format(int(text), ',')
    except ValueError:
        format_data = format(float(text), ',')
        if len(format_data.split('.')) >= 2:
            if dotdowndel:
                format_data = format_data.split('.')[0]
            elif dotdown8:
                if len(format_data.split('.')[1]) == 1:
                    format_data += '0000000'
                elif len(format_data.split('.')[1]) == 2:
                    format_data += '000000'
                elif len(format_data.split('.')[1]) == 3:
                    format_data += '00000'
                elif len(format_data.split('.')[1]) == 4:
                    format_data += '0000'
                elif len(format_data.split('.')[1]) == 5:
                    format_data += '000'
                elif len(format_data.split('.')[1]) == 6:
                    format_data += '00'
                elif len(format_data.split('.')[1]) == 7:
                    format_data += '0'
            elif len(format_data.split('.')[1]) == 1:
                format_data += '0'
    return format_data


def comma2int(t):
    if ' ' in t:
        t = t.split(' ')[1]
    if ',' in t:
        t = t.replace(',', '')
    return int(t)


def comma2float(t):
    if ' ' in t:
        t = t.split(' ')[1]
    if ',' in t:
        t = t.replace(',', '')
    return float(t)


def float2str1p6(seceonds):
    seceonds = str(seceonds)
    if len(seceonds.split('.')[1]) == 1:
        seceonds += '00000'
    elif len(seceonds.split('.')[1]) == 2:
        seceonds += '0000'
    elif len(seceonds.split('.')[1]) == 3:
        seceonds += '000'
    elif len(seceonds.split('.')[1]) == 4:
        seceonds += '00'
    elif len(seceonds.split('.')[1]) == 5:
        seceonds += '0'
    return seceonds

def parseDat(trcode, lines):
    lines = lines.split('\n')
    start = [i for i, x in enumerate(lines) if x.startswith('@START')]
    end = [i for i, x in enumerate(lines) if x.startswith('@END')]
    block = zip(start, end)
    enc_data = {'trcode': trcode, 'input': [], 'output': []}
    for start, end in block:
        block_data = lines[start - 1:end + 1]
        block_info = block_data[0]
        block_type = 'input' if 'INPUT' in block_info else 'output'
        record_line = block_data[1]
        tokens = record_line.split('_')[1].strip()
        record = tokens.split('=')[0]
        fields = block_data[2:-1]
        field_name = []
        for line in fields:
            field = line.split('=')[0].strip()
            field_name.append(field)
        fields = {record: field_name}
        enc_data['input'].append(fields) if block_type == 'input' else enc_data['output'].append(fields)
    return enc_data

def make_dir(path: str):
    if path == '.':
        return
    if not is_exist(path):
        os.mkdir(path)


def is_exist(path: str) -> bool:
    return os.path.exists(path)

def print_qsize(qlist : Queue):

    """
                    0        1       2       3      4       5       6       7        
        qlist = [tick0Q, tick1Q, tick2Q, tick3Q, hoga0Q, hoga1Q, hoga2Q, hoga3Q, 
                 broker0Q, broker1Q, broker2Q, broker3Q, viQ, save0Q, save1Q, save2Q, save3Q]
                    8         9         10        11     12     13      14     15       16
    """

    kospi = 'Kospi:{0:<6}'.format(qlist[0].qsize())
    kosdaq1 = 'Kosdaq1:{0:<6}'.format(qlist[1].qsize())
    kosdaq2 = 'Kosdaq2:{0:<6}'.format(qlist[2].qsize())

    save0 = 'Save0:{0:<4}'.format(qlist[13].qsize())
    save1 = 'Save1:{0:<4}'.format(qlist[14].qsize())
    save2 = 'Save2:{0:<4}'.format(qlist[15].qsize())
    Q_MSG = kospi + kosdaq1 + kosdaq2 + save0 + save1 + save2
    print(Q_MSG)
                              
    # print(f'Kospi:{qlist[0].qsize()} ', end='')
    # print(f'Kosdaq1:{qlist[1].qsize()} ', end='')
    # print(f'Kosdaq2:{qlist[2].qsize()} ', end='')
    # # print(f'Kosdaq_h[{qlist[3].qsize()}] ', end='')
    # print(f'Save1:{qlist[4].qsize()} ', end='')
    # print(f'Save2:{qlist[5].qsize()} ', end='')
    # print(f'Save3:{qlist[6].qsize()}')
