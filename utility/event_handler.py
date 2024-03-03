import time
from utility.constant import (
    TR_CODE_TICK_TYPE_MAP,
    TR_CODE_FIELDS_LIST_MAP,
)
# from datetime import datetime

class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("로그인 성공")
            XASessionEventHandler.login_state = 1
        else:
            print(f"로그인 실패 {msg}")


class XAQueryEventHandler:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandler.query_state = 1


class XARealEventHandler:
    def __init__(self):
        self.queue = None

    def handle_event(self, field_list: list) -> dict:
        values = {
            'system_time': time.time()
            # 'system_time': datetime.datetime.now()
        }
        for field in field_list:
            values[field] = self.GetFieldData("OutBlock", field)
        return values

    def OnReceiveRealData(self, tr_code):

        # tick_type = TR_CODE_TICK_TYPE_MAP.get(tr_code, None)
        tick_type = tr_code
        field_list = TR_CODE_FIELDS_LIST_MAP.get(tr_code, None)

        if tick_type is None or field_list is None:
            raise ValueError(f"Invalid TR code : {tr_code}")

        values = self.handle_event(field_list=field_list)
        data = (tick_type, values)
        self.queue.put(data)
