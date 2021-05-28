import csv
import ctypes
import os
import sys
import time
import traceback
from collections import OrderedDict, defaultdict
from datetime import datetime

import numpy as np
import pandas as pd
import schedule
import win32com.client
from slacker import Slacker

from config import config
from connect import connect
from holiday import is_holiday
from marketwatch import CpRpMarketWatch

indicators = {
    10: '외국계증권사창구첫매수',
    11: '외국계증권사창구첫매도',
    12: True,  # '외국인순매수',
    13: False,  # '외국인순매도',
    21: '전일거래량갱신',
    22: '최근5일거래량최고갱신',
    23: '최근5일매물대돌파',
    24: True,  # '최근60일매물대돌파',
    28: '최근5일첫상한가',
    29: False,  # '최근5일신고가갱신',
    30: '최근5일신저가갱신',
    31: '상한가직전',
    32: '하한가직전',
    41: '주가 5MA 상향돌파',
    42: '주가 5MA 하향돌파',
    43: '거래량 5MA 상향돌파',
    44: False,  # '주가데드크로스(5MA < 20MA)',
    45: True,  # '주가골든크로스(5MA > 20MA)',
    46: True,  # 'MACD 매수-Signal(9) 상향돌파',
    47: False,  # 'MACD 매도-Signal(9) 하향돌파',
    48: True,  # 'CCI 매수-기준선(-100) 상향돌파',
    49: False,  # 'CCI 매도-기준선(100) 하향돌파',
    50: True,  # 'Stochastic(10,5,5)매수- 기준선상향돌파',
    51: False,  # 'Stochastic(10,5,5)매도- 기준선하향돌파',
    52: True,  # 'Stochastic(10,5,5)매수- %config.K%D 교차',
    53: False,  # 'Stochastic(10,5,5)매도- %config.K%D 교차',
    54: True,  # 'Sonar 매수-Signal(9) 상향돌파',
    55: False,  # 'Sonar 매도-Signal(9) 하향돌파',
    56: True,  # 'Momentum 매수-기준선(100) 상향돌파',
    57: False,  # 'Momentum 매도-기준선(100) 하향돌파',
    58: True,  # 'RSI(14) 매수-Signal(9) 상향돌파',
    59: False,  # 'RSI(14) 매도-Signal(9) 하향돌파',
    60: True,  # 'Volume Oscillator 매수-Signal(9) 상향돌파',
    61: False,  # 'Volume Oscillator 매도-Signal(9) 하향돌파',
    62: True,  # 'Price roc 매수-Signal(9) 상향돌파',
    63: False,  # 'Price roc 매도-Signal(9) 하향돌파',
    64: True,  # '일목균형표매수-전환선 > 기준선상향교차',
    65: False,  # '일목균형표매도-전환선 < 기준선하향교차',
    66: True,  # '일목균형표매수-주가가선행스팬상향돌파',
    67: False,  # '일목균형표매도-주가가선행스팬하향돌파',
    68: True,  # '삼선전환도-양전환',
    69: False,  # '삼선전환도-음전환',
    70: True,  # '캔들패턴-상승반전형',
    71: False,  # '캔들패턴-하락반전형',
    81: '단기급락후 5MA 상향돌파',
    82: '주가이동평균밀집-5%이내',
    83: '눌림목재상승-20MA 지지'
}

slack = Slacker(config.token)

code_list = OrderedDict()
black_list = OrderedDict()
watch_data = {}
ohlc_list = {}
high_list = {}

pre_stock_message = ''
remark = ''


def print_message(*args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%Y-%m-%d %H:%M:%S]'), *args)
    print('----------------------------------------------------------------------------')


def slack_send_message(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    message = datetime.now().strftime('[%Y-%m-%d %H:%M:%S]\n') + message
    print(message)
    print('----------------------------------------------------------------------------')
    slack.chat.post_message('#stock', message)


def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print_message('check_creon_system() : admin user -> FAILED')
        return False

    # 연결 여부 체크
    if win32com.client.Dispatch('CpUtil.CpCybos').IsConnect == 0:
        print_message('check_creon_system() : connect to server -> FAILED')
        return False

    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if win32com.client.Dispatch('CpTrade.CpTdUtil').TradeInit(0) != 0:
        print_message('check_creon_system() : init trade -> FAILED')
        return False

    return True


def wait_for_request(check_type):
    """크레온 플러스 시스템 요청을 대기한다."""
    remain_count = cpStatus.GetLimitRemainCount(check_type)  # 0: 주문 관련 1: 시세 요청 관련 2: 실시간 요청 관련
    if 0 < remain_count:
        return

    remain_time = cpStatus.LimitRequestRemainTime
    print_message(f'대기시간: {remain_time / 1000:2.2f}초')
    time.sleep(remain_time / 1000)


def get_high_volume_code():
    """거래량 상위 종목 코드를 반환한다."""
    cpVolume.SetInputValue(0, ord('4'))  # 시장구분 4: 전체, 1: 거래소, 2: 코스닥
    cpVolume.SetInputValue(1, ord('V'))  # V: 거래량, A: 거래대금
    cpVolume.SetInputValue(2, ord('Y'))  # 관리 종목 제외 Y/N
    cpVolume.SetInputValue(3, ord('Y'))  # 우선주 제외 Y/N

    wait_for_request(1)
    cpVolume.BlockRequest()

    stop = cpVolume.GetHeaderValue(0)

    # ETF 먼저 담는다.
    for i in range(stop):
        if 200 <= len(code_list):
            break

        code = cpVolume.GetDataValue(1, i)  # 종목코드
        vol = cpVolume.GetDataValue(6, i)  # 거래량
        price = cpVolume.GetDataValue(3, i)  # 현재가
        percent = cpVolume.GetDataValue(5, i)  # 대비율
        name = cpVolume.GetDataValue(2, i)  # 종목명

        stockKind = cpCodeMgr.GetStockSectionKind(code)
        if stockKind == 10 or stockKind == 12:
            code_list[code] = (vol, price, percent, name)

    for i in range(stop):
        if 200 <= len(code_list):
            break

        code = cpVolume.GetDataValue(1, i)  # 종목코드
        vol = cpVolume.GetDataValue(6, i)  # 거래량
        price = cpVolume.GetDataValue(3, i)  # 현재가
        percent = cpVolume.GetDataValue(5, i)  # 대비율
        name = cpVolume.GetDataValue(2, i)  # 종목명

        # -15% 하락, 20% 상승 제외
        if percent < -15.0 or 20.0 < percent:
            continue

        if not code_list.get(code):
            code_list[code] = (vol, price, percent, name)


def get_biggest_moves_code():
    """상승 상위 종목 코드를 반환한다."""
    cpMoves.SetInputValue(0, ord('0'))  # 거래소 + 코스닥
    cpMoves.SetInputValue(1, ord('2'))  # 상승
    cpMoves.SetInputValue(2, ord('1'))  # 당일
    cpMoves.SetInputValue(3, 21)  # 전일 대비 상위 순
    cpMoves.SetInputValue(4, ord('1'))  # 관리 종목 제외
    cpMoves.SetInputValue(5, ord('0'))  # 거래량 전체
    cpMoves.SetInputValue(6, ord('0'))  # '표시 항목 선택 - '0': 시가대비
    cpMoves.SetInputValue(7, 0)  # 등락율 시작
    cpMoves.SetInputValue(8, 20)  # 등락율 끝, 20% 상승 제외

    wait_for_request(1)
    cpMoves.BlockRequest()

    for i in range(cpMoves.GetHeaderValue(0)):
        if 200 <= len(code_list):
            break

        # 상장 주식수 20억 이상만 담는다.
        code = cpMoves.GetDataValue(0, i)  # 코드
        vol = cpMoves.GetDataValue(6, i)  # 거래량
        price = cpMoves.GetDataValue(2, i)  # 현재가
        percent = cpMoves.GetDataValue(4, i)  # 대비율
        name = cpMoves.GetDataValue(1, i)  # 종목명

        if not cpCodeMgr.IsBigListingStock(code):
            continue

        # -15% 하락 제외
        if percent < -15.0:
            continue

        code_list[code] = (vol, price, percent, name)


def get_market_cap(codes):
    """시가총액 순으로 종목 코드를 변환한다."""
    cpMarketEye.SetInputValue(0, [0, 4, 20])  # 0: 종목코드 4: 현재가 20: 상장주식수
    cpMarketEye.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트

    wait_for_request(1)
    cpMarketEye.BlockRequest()

    market_caps = OrderedDict()
    for i in range(cpMarketEye.GetHeaderValue(2)):
        code = cpMarketEye.GetDataValue(0, i)  # 코드
        close = cpMarketEye.GetDataValue(1, i)  # 종가
        listed_stock = cpMarketEye.GetDataValue(2, i)  # 상장주식수

        market_cap = close * listed_stock
        if cpCodeMgr.IsBigListingStock(code):
            market_cap *= 1000
        market_cap = market_cap // 100000000
        market_caps[code] = market_cap

    return market_caps


def sort_code_list(market_caps: OrderedDict):
    """시가총액 상위 순으로 종목 코드를 정렬한다."""
    if len(code_list) <= 0 or len(market_caps) <= 0:
        return

    for (code, (vol, price, percent, name)), (code2, (market_cap)) in zip(code_list.items(), market_caps.items()):
        if code == code2:
            code_list[code] = (vol, price, percent, name, market_cap)

    temp = OrderedDict(sorted(code_list.items(), key=lambda x: x[1][4], reverse=True)[:config.code_limit])
    code_list.clear()
    code_list.update(temp)


def print_code_list():
    """종목 코드를 출력한다."""
    message = '\n코드\t거래량      시가총액(억 원)  현재가(원)  대비율\t종목명\n'
    for code, item in code_list.items():
        vol, price, percent, name, market_cap = item
        message += f'{code}\t{vol:11,}\t{market_cap:10,}\t{price:7,}\t{percent:>6.02f}\t{name:20}\n'
    print_message(message)


def get_code_list():
    """종목 코드를 가져온다."""
    code_list.clear()
    get_high_volume_code()
    get_biggest_moves_code()
    market_caps = get_market_cap(list(code_list.keys()))
    sort_code_list(market_caps)
    print_code_list()


def get_watch_data():
    """특징주 포착을 수신한다."""
    wait_for_request(2)
    cpRpMarketWatch.Request('*', watch_data)


def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)  # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째

    wait_for_request(0)
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9)  # 증거금 100% 주문 가능 금액


def get_stock_balance(code=''):
    """인자로 받은 종목의 종목명과 수량, 장부가를 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpStockBalance.SetInputValue(0, acc)  # 계좌번호
    cpStockBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpStockBalance.SetInputValue(2, 50)  # 요청 건수(최대 50)

    wait_for_request(0)
    cpStockBalance.BlockRequest()

    stock_balance = {}
    for i in range(cpStockBalance.GetHeaderValue(7)):
        stock_code = cpStockBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpStockBalance.GetDataValue(0, i)  # 종목명
        stock_qty = cpStockBalance.GetDataValue(15, i)  # 수량
        stock_price = cpStockBalance.GetDataValue(17, i)  # 장부가
        eval_percentage = cpStockBalance.GetDataValue(11, i)  # 평가손익

        if stock_code == code:
            return stock_name, stock_qty, stock_price

        stock_balance[stock_code] = {
            'name': stock_name,
            'shares': stock_qty,
            'price': stock_price,
            'percentage': eval_percentage
        }
    if code != '':
        stock_name = cpStockCode.CodeToName(code)
        return stock_name, 0, 0
    else:
        return stock_balance


def print_stock_balance(stock_balance):
    """보유 종목을 출력한다."""
    global pre_stock_message

    message = '주식잔고\n'
    message += '코드\t수량  대비율\t최고익\t장부가\t종목명\n'
    for code in stock_balance:
        stock = stock_balance[code]
        shares = stock['shares']
        percentage = stock['percentage']
        high = high_list[code] if high_list.get(code) else 0.0
        price = int(stock['price'])
        name = stock['name']
        message += f'{code}\t{shares:>5,}\t{percentage:>5.02f}%\t{high:>5.02f}%\t{price:>,}\t{name}\n'

    # 이전 잔고메세지랑 다를때만 출력
    if pre_stock_message != message:
        print_message(message)
    pre_stock_message = message


def get_balance():
    """수익률, 잔량평가손익, 매도실현손익을 파이썬 셸과 동시에 슬랙으로 출력한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)  # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째

    wait_for_request(0)
    cpBalance.BlockRequest()

    yield_rate = cpBalance.GetHeaderValue(3)
    yield_rate = 0.0 if yield_rate == '' else float(yield_rate)
    slack_send_message(f'수익률: `{yield_rate:>2.2f}`%')


def get_current_stock(code):
    """인자로 받은 종목의 현재가, 고가, 저가를 반환한다."""
    cpStockMst.SetInputValue(0, code)  # 종목코드에 대한 가격 정보

    wait_for_request(1)
    cpStockMst.BlockRequest()

    current_price = cpStockMst.GetHeaderValue(11)  # 현재가
    high = cpStockMst.GetHeaderValue(14)  # 고가
    low = cpStockMst.GetHeaderValue(15)  # 저가

    return current_price, high, low


def has_enough_cash(current_price):
    """매수할 때 100% 증거금이 충분한지 조회한다."""
    total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
    buy_amount = int(config.buy_amount)
    shares = int(buy_amount // current_price)
    if total_cash < current_price * shares:
        shares = int(total_cash // current_price)
        if shares == 0:
            print(f'100% 증거금 {current_price - total_cash:,}원 부족')
            return False, 0

    if shares == 0:
        print(f'종목당 주문가능금액 {current_price - config.buy_amount:,}원 부족')
        return False, 0

    return True, shares


def sell_stock(code, name, shares, percentage):
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션

        # 최유리 IOC 매도 주문 설정
        cpOrder.SetInputValue(0, "1")  # 1:매도, 2:매수
        cpOrder.SetInputValue(1, acc)  # 계좌번호
        cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
        cpOrder.SetInputValue(3, code)  # 종목코드
        cpOrder.SetInputValue(4, shares)  # 매도수량
        cpOrder.SetInputValue(7, "1")  # 조건 0:기본, 1:IOC, 2:FOK
        cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선

        # 최유리 IOC 매도 주문 요청
        ret = cpOrder.BlockRequest()
        if ret == 4:
            print_message('주의: 연속 주문 제한')
            wait_for_request(0)
            cpOrder.BlockRequest()
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message(f"`sell({code}) -> exception! " + str(e) + "`")


def get_transaction_history(code=""):
    """ 금일 계좌별 주문/체결 내역 조회 데이터를 요청하고 수신한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
    cpTrade.SetInputValue(0, acc)  # 계좌번호
    cpTrade.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpTrade.SetInputValue(2, code)  # 종목코드[default:""] - 생략 시 전종목에 대해서 조회가됨

    wait_for_request(0)
    cpTrade.BlockRequest()

    history = defaultdict(list)
    for i in range(cpTrade.GetHeaderValue(6)):
        code = cpTrade.GetDataValue(3, i)  # 종목코드
        name = cpTrade.GetDataValue(4, i)  # 종목이름
        quantity = cpTrade.GetDataValue(9, i)  # 총체결수량
        price = cpTrade.GetDataValue(11, i)  # 체결단가
        state = cpTrade.GetDataValue(13, i)  # 정정취소구분내용
        order = cpTrade.GetDataValue(35, i)  # 매매구분코드 1: 매도, 2: 매수
        history[code].append({
            'name': name,
            'quantity': quantity,
            'price': price,
            'state': state,
            'order': order
        })

    return history


def buy_stock(code, name, shares, current_price):
    """인자로 받은 종목을 최유리 지정가 IOC 조건으로 매수한다."""
    try:
        # 매수 완료 종목이면 더 이상 안 사도록 리턴
        stock_balance = get_stock_balance()
        if stock_balance.get(code):
            print_message(f'{code} {name}\n'
                          f'이미 매수한 종목이므로 더 이상 구매하지 않습니다.')
            return

        # 매수 완료 종목 개수가 매수할 종목 수 이상이면 리턴
        if config.target_buy_count <= len(stock_balance):
            print_message(f'매수한 종목 수: {len(stock_balance)}\n'
                          f'더 이상 구매하지 않습니다.')
            return

        # 금일 계좌에 체결내역이 있을 경우 구매하지 않음
        history = get_transaction_history(code)
        if code in history.keys():
            # 체결내역이 정상주문에 매수일 경우에
            for item in history[code]:
                if '정상주문' == item["state"]:
                    print_message(f'거래 내역에 해당 종목이 있습니다.\n'
                                  f'{code} {name}\n'
                                  f'체결단가: {item["price"]:,}\n'
                                  f'현재가: {current_price:,}')
                    return

        read_blacklist()
        if code in black_list.keys():
            slack_send_message(f'블랙리스트에 해당 종목({black_list[code]})이 있습니다.')
            return

        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션

        # 최유리 IOC 매수 주문 설정
        cpOrder.SetInputValue(0, "2")  # 2: 매수
        cpOrder.SetInputValue(1, acc)  # 계좌번호
        cpOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        cpOrder.SetInputValue(3, code)  # 종목코드
        cpOrder.SetInputValue(4, shares)  # 매수할 수량
        cpOrder.SetInputValue(7, "1")  # 주문조건 0:기본, 1:IOC, 2:FOK
        cpOrder.SetInputValue(8, "12")  # 주문호가 1:보통, 3:시장가 5:조건부, 12:최유리, 13:최우선

        # 매수 주문 요청
        ret = cpOrder.BlockRequest()

        if ret == 4:
            print_message('주의: 연속 주문 제한')
            wait_for_request(0)
            cpOrder.BlockRequest()
        stock_name, stock_qty, stock_price = get_stock_balance(code)
        if shares <= stock_qty:
            slack_send_message(f'{name} {stock_price:,.1f}원 {shares}주 매수')
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`buy_stock(" + str(code) + ") -> exception! " + str(e) + "`")


def get_ohlc(code, window):
    """인자로 받은 종목의 OHLC 가격 정보를 shares 개수만큼 반환한다."""
    path = './ohlc'
    os.makedirs(path, exist_ok=True)

    today = datetime.now().strftime('%Y-%m-%d')
    file_path = path + '/' + code + '-' + today + '.csv'

    if os.path.isfile(file_path):
        df = pd.read_csv(file_path)
    else:
        columns = ['open', 'high', 'low', 'close']
        index, rows = [], []
        cpOhlc.SetInputValue(0, code)  # 종목코드
        cpOhlc.SetInputValue(1, ord('2'))  # 1:기간, 2:개수
        cpOhlc.SetInputValue(4, window)  # 요청개수
        cpOhlc.SetInputValue(5, [0, 1, 2, 3, 4, 5])  # 0:날짜, 1:시간, 2~5:시가,고가,저가,종가
        cpOhlc.SetInputValue(6, ord('D'))  # D:일단위, m:분단위
        cpOhlc.SetInputValue(9, ord('1'))  # 0:무수정주가, 1:수정주가

        wait_for_request(1)
        cpOhlc.BlockRequest()
        for i in range(cpOhlc.GetHeaderValue(3)):  # 3:수신개수
            index.append(cpOhlc.GetDataValue(0, i))
            rows.append([cpOhlc.GetDataValue(2, i),
                         cpOhlc.GetDataValue(3, i),
                         cpOhlc.GetDataValue(4, i),
                         cpOhlc.GetDataValue(5, i)])

        df = pd.DataFrame(rows, index=index, columns=columns)

        df.to_csv(file_path, index=True)

    return df


def get_ror(ohlc, k):
    ohlc['range'] = (ohlc['high'] - ohlc['low']) * k
    ohlc['target'] = ohlc['open'] + ohlc['range'].shift(1)

    ohlc['ror'] = np.where(ohlc['high'] > ohlc['target'], ohlc['close'] / ohlc['target'], 1)

    return max(ohlc['ror'].cumprod())


def get_k(ohlc):
    config.K = 0.1
    max_ror = get_ror(ohlc, 0.1)

    for k in [p / 10 for p in range(2, 10)]:
        ror = get_ror(ohlc, k)
        if max_ror < ror:
            max_ror = ror
            config.K = k


def get_target_price(ohlc):
    """매수 목표가를 반환한다."""
    target_price = 0

    try:
        lastday_key = ohlc.iloc[0].name
        lastday = ohlc.loc[lastday_key]
        today_open = lastday['close']
        lastday_high = lastday['high']
        lastday_low = lastday['low']
        get_k(ohlc)
        target_price = today_open + (lastday_high - lastday_low) * config.K
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`get_target_price() -> exception! " + str(e) + "`")

    return target_price


def get_predicted_price(code):
    """예상 목표가를 반환한다."""
    predicted_price = 0

    try:
        file_path = './predict/' + code + '.csv'
        if os.path.isfile(file_path):
            with open(file_path, 'r', encoding="utf-8") as f:
                csv_reader = csv.reader(f)
                for row in csv_reader:
                    predicted_price = float(row[1])
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`get_target_price() -> exception! " + str(e) + "`")

    return predicted_price


def get_movingaverage(ohlc, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        closes = ohlc['close'].sort_index()
        ma = closes.rolling(window=window, min_periods=1).mean()
        lastday_key = ohlc.iloc[0].name
        return ma.loc[lastday_key]
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message('`get_movingaverage(' + str(window) + ') -> exception! ' + str(e) + "`")
        return None


def buy_watch_data():
    """주요 신호 포착될 때 매수한다."""
    try:
        global remark

        for code in watch_data.keys():
            item = watch_data[code]
            if item['indicator'] in indicators \
                    and indicators[item['indicator']] is True:
                if not cpCodeMgr.IsBigListingStock(code):
                    continue

                name = cpCodeMgr.CodeToName(code)
                current_price, _, _ = get_current_stock(code)
                enough, shares = has_enough_cash(current_price)
                if enough:
                    message = f'[{item["time"]}] {code} {name}, {item["remark"]}'
                    if remark != message:
                        slack_send_message(message)
                    remark = message
                    buy_stock(code, name, shares, current_price)
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`buy_watch_data() -> exception! " + str(e) + "`")


def sell_watch_data():
    """주요 신호 포착될 때 매도한다."""
    try:
        global remark

        stock_balance = get_stock_balance()

        print_stock_balance(stock_balance)

        for code in stock_balance:
            stock = stock_balance[code]
            percentage = stock['percentage']
            shares = stock['shares']
            if code in watch_data.keys():
                item = watch_data[code]
                if item['indicator'] in indicators \
                        and indicators[item['indicator']] is False:
                    name = cpCodeMgr.CodeToName(code)
                    message = f'[{item["time"]}] {code} {name}, {item["remark"]}'
                    if remark != message:
                        slack_send_message(remark)
                    remark = message
                    sell_stock(code, name, shares, percentage)
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`sell_watch_data() -> exception! " + str(e) + "`")


def read_blacklist():
    try:
        if os.path.isfile('blacklist.csv'):
            with open('blacklist.csv', 'r', encoding="utf-8") as f:
                csv_reader = csv.reader(f)
                for row in csv_reader:
                    if row:
                        black_list[row[0]] = [row[1], row[2]]
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`read_blacklist() -> exception! " + str(e) + "`")


def write_blacklist(code, name, percentage):
    try:
        with open('blacklist.csv', 'a', encoding="utf-8", newline='\n') as f:
            csv_writer = csv.writer(f)
            row = [code, name, percentage]
            csv_writer.writerow(row)
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`write_blacklist() -> exception! " + str(e) + "`")


def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        stock_balance = get_stock_balance()

        print_stock_balance(stock_balance)

        for code in stock_balance:
            stock = stock_balance[code]
            name = stock['name']
            shares = stock['shares']
            price = stock['price']
            percentage = stock['percentage']

            predicted_price = get_predicted_price(code)  # 예상 목표가

            if 0 < percentage and predicted_price and predicted_price < price:
                sell_stock(code, name, shares, percentage)
                slack_send_message(f'{name} {shares}주 매도\n'
                                   f'손익: `{percentage:2.2f}`\n'
                                   f'예상가: {predicted_price:2.2f}\n'
                                   f'현재가: {price:2.2f}')
            else:
                if code not in high_list.keys():
                    high_list[code] = percentage
                else:
                    high_list[code] = max(percentage, high_list[code])

                if config.profit_rate <= percentage:
                    if percentage < high_list[code]:
                        sell_stock(code, name, shares, percentage)
                        slack_send_message(f'{name} {shares}주 매도 (손익: `{percentage:2.2f}`)')

                if percentage <= config.loss_rate:
                    sell_stock(code, name, shares, percentage)
                    write_blacklist(code, name, percentage)
                    slack_send_message(f'{name} {shares}주 매도 (손익: `{percentage:2.2f}`)')
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`sell_all() -> exception! " + str(e) + "`")


def sell_all_and_buy_code_list():
    """종목 코드의 목표가 보다 현재가가 클 때 매수한다."""
    try:
        for code in code_list.keys():
            sell_all()

            get_curr(code)

            if code not in ohlc_list.keys():
                ohlc = get_ohlc(code, 10)
                ohlc_list[code] = ohlc
            else:
                ohlc = ohlc_list[code]
            target_price = get_target_price(ohlc)  # 매수 목표가
            predicted_price = get_predicted_price(code)  # 예상 목표가
            ma5_price = get_movingaverage(ohlc, 5)  # 5일 이동평균가
            ma10_price = get_movingaverage(ohlc, 10)  # 10일 이동평균가
            current_price, high, low = get_current_stock(code)
            name = code_list[code][3]
            # print(name, current_price, target_price, high, ma5_price, ma10_price)
            # 매수 목표가, 5일 이동평균가, 10일 이동평균가 보다 현재가가 클 때 매수
            if target_price < current_price < predicted_price \
                    and current_price + current_price * (config.profit_rate / 100) < predicted_price \
                    and ma5_price < current_price \
                    and ma10_price < current_price:
                print(datetime.now().strftime('[%Y-%m-%d %H:%M:%S]'))
                print(f'{name}')
                print(f'현재가: {current_price:,}원')
                print(f'목표가: {int(target_price):,}원')
                print(f'예상가: {int(predicted_price):,}원')
                print(f'MA05  : {int(ma5_price):,}원')
                print(f'MA10  : {int(ma10_price):,}원')
                enough, shares = has_enough_cash(current_price)
                if enough:
                    buy_stock(code, name, shares, current_price)
                print('----------------------------------------------------------------------------')
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`buy_code_list() -> exception! " + str(e) + "`")


def getHMTFromTime(str_time):
    from datetime import time
    hh, mm = divmod(str_time, 10000)
    mm, tt = divmod(mm, 100)
    return time(hh, mm, tt).strftime("%H:%M:%S")


def get_curr(code):
    path = './curr'
    os.makedirs(path, exist_ok=True)

    today = datetime.now().strftime('%Y-%m-%d')
    file_path = path + '/' + code + '-' + today + '.csv'

    # 현재가 통신
    cpStockBid.SetInputValue(0, code)
    cpStockBid.SetInputValue(2, 80)  # 요청개수 (최대 80)
    cpStockBid.SetInputValue(3, ord('C'))  # C 체결가 비교 방식 H 호가 비교방식

    cpStockBid.BlockRequest()

    if cpStockBid.GetDibStatus() != 0:
        print("통신상태", cpStockBid.GetDibStatus(), cpStockBid.GetDibMsg1())
        return False

    columns = ['ds', 'y']
    rows = []
    for i in range(cpStockBid.GetHeaderValue(2)):
        rows.append([
            today + ' ' + getHMTFromTime(cpStockBid.GetDataValue(9, i)),
            cpStockBid.GetDataValue(4, i)
        ])

    df = pd.DataFrame(rows, columns=columns)

    if os.path.isfile(file_path):
        df = df.append(pd.read_csv(file_path), ignore_index=True)

    df.to_csv(file_path, index=False)

    return df


def auto_trade():
    """자동 매도, 매수, 종료한다."""
    t_now = datetime.now()
    t_start = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
    t_exit = t_now.replace(hour=15, minute=30, second=0, microsecond=0)

    # AM 09:00 ~ PM 15:30 : 매도 & 매수
    if t_start < t_now < t_exit:
        sell_watch_data()
        buy_watch_data()
        sell_all_and_buy_code_list()

    # PM 15:30 ~ :프로그램 종료
    if t_exit < t_now:
        slack_send_message('`장 마감`')
        time.sleep(1)
        get_balance()
        os.system('taskkill /IM CpStart* /F /T')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        sys.exit(0)


if __name__ == '__main__':
    try:
        if is_holiday():  # 휴일
            print_message('Today is holiday')
            sys.exit(0)

        if not check_creon_system():
            connect()

        # 크레온 플러스 공통 OBJECT
        cpBalance = win32com.client.Dispatch("CpTrade.CpTd6032")
        cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
        cpStockBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
        cpStockCode = win32com.client.Dispatch('CpUtil.CpStockCode')
        cpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        cpVolume = win32com.client.Dispatch("CpSysDib.CpSvr7049")
        cpMoves = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        cpMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
        cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
        cpStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
        cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
        cpTrade = win32com.client.Dispatch('CpTrade.CpTd5341')
        cpStockBid = win32com.client.Dispatch("Dscbo1.StockBid")
        cpRpMarketWatch = CpRpMarketWatch()

        print_message('시작 시간')

        get_watch_data()
        get_code_list()
        schedule.every(15).seconds.do(get_watch_data)
        schedule.every(15).seconds.do(get_code_list)

        schedule.every(1).seconds.do(auto_trade)

        while True:
            schedule.run_pending()
    except Exception as ex:
        traceback.print_exc(file=sys.stdout)
        print_message('`main -> exception! ' + str(ex) + '`')
