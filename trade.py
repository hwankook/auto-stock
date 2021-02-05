import ctypes
import sys
import time
import traceback
from collections import OrderedDict
from datetime import datetime

import pandas as pd
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
    remain_count = cpStatus.GetLimitRemainCount(check_type)  # 0: 주문 관련 1: 시세 요청 관련 2: 실시간 요청 관련
    if 0 < remain_count:
        return

    remain_time = cpStatus.LimitRequestRemainTime
    print(f'대기시간: {remain_time / 1000:2.2f}초')
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

        if not cpCodeMgr.IsBigListingStock(code):
            continue

        # 상장 주식수 20억 이상만 담는다.
        code = cpMoves.GetDataValue(0, i)  # 코드
        vol = cpMoves.GetDataValue(6, i)  # 거래량
        price = cpMoves.GetDataValue(2, i)  # 현재가
        percent = cpMoves.GetDataValue(4, i)  # 대비율
        name = cpMoves.GetDataValue(1, i)  # 종목명

        # -15% 하락 제외
        if percent < -15.0:
            continue

        code_list[code] = (vol, price, percent, name)


def get_market_cap(codes):
    """시가총액 순으로 종목코드를 변환한다."""
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
    for (code, (vol, price, percent, name)), (code2, (market_cap)) in zip(code_list.items(), market_caps.items()):
        if code == code2:
            code_list[code] = (vol, price, percent, name, market_cap)
    return OrderedDict(sorted(code_list.items(), key=lambda x: x[1][4], reverse=True)[:config.code_limit])


def print_code_list():
    """종목 코드를 출력한다."""
    message = '\n코드\t거래량      시가총액(억 원)  현재가(원)  대비율\t종목명\n'
    for code, item in code_list.items():
        vol, price, percent, name, market_cap = item
        message += f'{code}\t{vol:11,}\t{market_cap:10,}\t{price:7,}\t{percent:>6.02f}\t{name:20}\n'
    print_message(message)


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


def get_current_price(code):
    """인자로 받은 종목의 현재가, 매도호가, 매도호가 잔량, 매수호가, 매수호가 잔량을 반환한다."""
    cpStockMst.SetInputValue(0, code)  # 종목코드에 대한 가격 정보

    wait_for_request(1)
    cpStockMst.BlockRequest()

    current_price = cpStockMst.GetHeaderValue(11)  # 현재가

    return current_price


def has_enough_cash(code, name, current_price=-1):
    total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
    if current_price == -1:
        current_price = get_current_price(code)
    buy_amount = int(config.buy_amount)
    shares = int(buy_amount // current_price)
    if total_cash < current_price * shares:
        shares = int(total_cash // current_price)
        if shares == 0:
            print_message(f'{name} {current_price:,}\n'
                          f'100% 증거금 {current_price - total_cash:,}원 부족')
            return False, 0

    if shares == 0:
        print_message(f'{name} {current_price:,}\n'
                      f'종목당 주문가능금액 {current_price - config.buy_amount:,}원 부족')
        return False, 0

    return True, shares


def get_stock_balance(code=''):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
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
            return stock_name, stock_qty

        stock_balance[stock_code] = {
            'name': stock_name,
            'shares': stock_qty,
            'price': stock_price,
            'percentage': eval_percentage
        }
    if code != '':
        stock_name = cpStockCode.CodeToName(code)
        return stock_name, 0
    else:
        return stock_balance


def print_stock_balance(stock_balance):
    """보유 종목을 출력한다."""
    if 0 < len(stock_balance):
        message = '주식잔고\n'
        message += '코드\t수량  대비율\t장부가\t종목명\n'
        for code, stock in stock_balance.items():
            shares = stock['shares']
            percentage = stock['percentage']
            price = int(stock['price'])
            name = stock['name']
            message += f'{code}\t{shares:>5,}\t{percentage:>5.02f}%\t{price:>,}\t{name}\n'
        print_message(message)


def get_ohlc(code, window):
    """인자로 받은 종목의 OHLC 가격 정보를 shares 개수만큼 반환한다."""
    columns = ['open', 'high', 'low', 'close']
    index, rows = [], []
    cpOhlc.SetInputValue(0, code)  # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))  # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, window)  # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5])  # 0:날짜, 2~5:시가,고가,저가,종가
    cpOhlc.SetInputValue(6, ord('D'))  # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))  # 0:무수정주가, 1:수정주가

    wait_for_request(1)
    cpOhlc.BlockRequest()
    for i in range(cpOhlc.GetHeaderValue(3)):  # 3:수신개수
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i),
                     cpOhlc.GetDataValue(2, i),
                     cpOhlc.GetDataValue(3, i),
                     cpOhlc.GetDataValue(4, i)])
    df = pd.DataFrame(rows, index=index, columns=columns)

    return df


def get_target_price_to_buy(ohlc):
    """매수 목표가를 반환한다."""
    try:
        if len(ohlc) <= 1:
            lastday = ohlc.iloc[0]
            today_open = lastday['close']
        else:
            str_today = datetime.now().strftime('%Y%m%d')
            if str_today == str(ohlc.iloc[0].name):
                lastday = ohlc.iloc[1]
                today_open = ohlc.iloc[0]['open']
            else:
                lastday = ohlc.iloc[0]
                today_open = lastday['close']
        lastday_high = lastday['high']
        lastday_low = lastday['low']
        target_price = today_open + (lastday_high - lastday_low) * config.K
        return target_price
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`get_target_price() -> exception! " + str(e) + "`")
        return None


def get_movingaverage(ohlc, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        closes = ohlc['close'].sort_index()
        if len(ohlc) < window:
            temp = ohlc.iloc[::-1]
            lastday = temp.iloc[-1].name
            ma = closes.rolling(window=len(ohlc)).mean()
            return ma.loc[lastday]

        str_today = datetime.now().strftime('%Y%m%d')
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message('`get_movingaverage(' + str(window) + ') -> exception! ' + str(e) + "`")
        return None


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
        slack_send_message(f'{name} {shares} 주 매도 (손익: `{percentage:2.2f}`) -> returned {ret}')
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message(f"`sell({code}) -> exception! " + str(e) + "`")


def sell_all(listWatchData):
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        stock_balance = get_stock_balance()
        print_stock_balance(stock_balance)

        for code, stock in stock_balance.items():
            percentage = stock['percentage']
            name = stock['name']
            shares = stock['shares']

            # 손익 +1.3%, -2% 매도
            if config.profit_rate <= percentage or percentage <= config.loss_rate:
                # 시가총액 10조 이상이면 +10%, -5% 매도
                market_caps = get_market_cap(code)
                if 100000 < market_caps[code]:
                    if 10.0 <= percentage or percentage <= -5.0:
                        sell_stock(code, name, shares, percentage)
                        continue

                sell_stock(code, name, shares, percentage)
                continue

            # 주요 신호 포착될 떄 매도
            if code in listWatchData.keys():
                item = listWatchData[code]
                if item['indicator'] in indicators \
                        and indicators[item['indicator']] is False:
                    name = cpCodeMgr.CodeToName(code)
                    slack_send_message(f'[{item["time"]}] {code} {name}, {item["remark"]}')
                    sell_stock(code, name, shares, percentage)
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`sell_all() -> exception! " + str(e) + "`")


def buy_stock(code, name, shares):
    """인자로 받은 종목을 최유리 지정가 IOC 조건으로 매수한다."""
    try:
        # 매수 완료 종목이면 더 이상 안 사도록 리턴
        stock_balance = get_stock_balance()
        if stock_balance.get(code):
            print_message(f'{code} {name}\n'
                          f'이미 매수한 종목이므로 더 이상 구매하지 않습니다.')
            return

        # 매수 완료 종목 개수가 매수할 종목 수 이상이면 리턴
        if config.target_buy_count < len(stock_balance):
            print_message(f'매수한 종목 수: {len(stock_balance)}\n'
                          f'더 이상 구매하지 않습니다.')
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
        stock_name, stock_qty = get_stock_balance(code)
        if shares <= stock_qty:
            slack_send_message(f'{name} {shares}주 매수 -> returned {ret}')
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`buy_etf(" + str(code) + ") -> exception! " + str(e) + "`")


def buy_all(listWatchData):
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        # 주요 신호 포착될 때
        for code in listWatchData.keys():
            item = listWatchData[code]
            if item['indicator'] in indicators \
                    and indicators[item['indicator']] is True:
                name = cpCodeMgr.CodeToName(code)
                slack_send_message(f'[{item["time"]}] {code} {name}, {item["remark"]}')
                enough, shares = has_enough_cash(code, name)
                if enough:
                    buy_stock(code, name, shares)

        # 매수 목표가, 5일 이동평균가, 10일 이동평균가 보다 현재가가 클 때 매수
        for code in code_list.keys():
            ohlc = get_ohlc(code, 10)
            target_price = get_target_price_to_buy(ohlc)  # 매수 목표가
            ma5_price = get_movingaverage(ohlc, 5)  # 5일 이동평균가
            ma10_price = get_movingaverage(ohlc, 10)  # 10일 이동평균가
            current_price = int(code_list[code][1])
            if target_price < current_price \
                    and ma5_price < current_price \
                    and ma10_price < current_price:
                name = code_list[code][3]
                enough, shares = has_enough_cash(code, name, current_price)
                if enough:
                    buy_stock(code, name, shares)
    except Exception as e:
        traceback.print_exc(file=sys.stdout)
        slack_send_message("`sell_all() -> exception! " + str(e) + "`")


def auto_trade():
    while True:
        get_high_volume_code()
        get_biggest_moves_code()
        market_caps = get_market_cap(list(code_list.keys()))
        temp = sort_code_list(market_caps)
        code_list.clear()
        code_list.update(temp)
        print_code_list()

        t_now = datetime.now()
        t_start = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
        t_exit = t_now.replace(hour=15, minute=30, second=0, microsecond=0)
        if t_start < t_now < t_exit:  # AM 09:05 ~ PM 03:15 : 매수 & 매도
            total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
            print_message(f'100% 증거금 주문가능금액: {total_cash:,}')

            listWatchData = {}
            wait_for_request(2)
            cpRpMarketWatch.Request('*', listWatchData)

            sell_all(listWatchData)

            buy_all(listWatchData)

        if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
            slack_send_message('`장 마감`')
            time.sleep(1)
            get_balance()
            sys.exit(0)

        code_list.clear()

        time.sleep(15)


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
        cpRpMarketWatch = CpRpMarketWatch()

        print_message('시작 시간')

        auto_trade()
    except Exception as ex:
        traceback.print_exc(file=sys.stdout)
        print_message('`main -> exception! ' + str(ex) + '`')
