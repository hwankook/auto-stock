from collections import OrderedDict

import win32com.client

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

indicators = {
    10: '외국계증권사창구첫매수',
    11: '외국계증권사창구첫매도',
    12: '외국인순매수',
    13: '외국인순매도',
    21: '전일거래량갱신',
    22: '최근5일거래량최고갱신',
    23: '최근5일매물대돌파',
    24: '최근60일매물대돌파',
    28: '최근5일첫상한가',
    29: '최근5일신고가갱신',
    30: '최근5일신저가갱신',
    31: '상한가직전',
    32: '하한가직전',
    41: '주가 5MA 상향돌파',
    42: '주가 5MA 하향돌파',
    43: '거래량 5MA 상향돌파',
    44: '주가데드크로스(5MA < 20MA)',
    45: '주가골든크로스(5MA > 20MA)',
    46: 'MACD 매수-Signal(9) 상향돌파',
    47: 'MACD 매도-Signal(9) 하향돌파',
    48: 'CCI 매수-기준선(-100) 상향돌파',
    49: 'CCI 매도-기준선(100) 하향돌파',
    50: 'Stochastic(10,5,5)매수- 기준선상향돌파',
    51: 'Stochastic(10,5,5)매도- 기준선하향돌파',
    52: 'Stochastic(10,5,5)매수- %K%D 교차',
    53: 'Stochastic(10,5,5)매도- %K%D 교차',
    54: 'Sonar 매수-Signal(9) 상향돌파',
    55: 'Sonar 매도-Signal(9) 하향돌파',
    56: 'Momentum 매수-기준선(100) 상향돌파',
    57: 'Momentum 매도-기준선(100) 하향돌파',
    58: 'RSI(14) 매수-Signal(9) 상향돌파',
    59: 'RSI(14) 매도-Signal(9) 하향돌파',
    60: 'Volume Oscillator 매수-Signal(9) 상향돌파',
    61: 'Volume Oscillator 매도-Signal(9) 하향돌파',
    62: 'Price roc 매수-Signal(9) 상향돌파',
    63: 'Price roc 매도-Signal(9) 하향돌파',
    64: '일목균형표매수-전환선 > 기준선상향교차',
    65: '일목균형표매도-전환선 < 기준선하향교차',
    66: '일목균형표매수-주가가선행스팬상향돌파',
    67: '일목균형표매도-주가가선행스팬하향돌파',
    68: '삼선전환도-양전환',
    69: '삼선전환도-음전환',
    70: '캔들패턴-상승반전형',
    71: '캔들패턴-하락반전형',
    81: '단기급락후 5MA 상향돌파',
    82: '주가이동평균밀집-5%이내',
    83: '눌림목재상승-20MA 지지'
}


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, listWatchData: OrderedDict):
        self.client = client  # CP 실시간 통신 object
        self.listWatchData = listWatchData

    def OnReceived(self):
        # 실시간 처리 - marketwatch : 특이 신호(차트, 외국인 순매수 등)
        code = self.client.GetHeaderValue(0)

        for i in range(self.client.GetHeaderValue(2)):
            item = {
                'code': code,
                'indicator': ''
            }
            indicator = self.client.GetDataValue(2, i)
            if indicator in indicators:
                update = self.client.GetDataValue(1, i)
                if update != ord('c'):
                    item['indicator'] = indicator
                    item['remark'] = indicators[indicator]
            self.listWatchData[code] = item


class CpPublish:
    def __init__(self, serviceID):
        self.obj = win32com.client.Dispatch(serviceID)

    def Unsubscribe(self):
        self.obj.Unsubscribe()

    def Subscribe(self, code, listWatchData):
        if 0 < len(code):
            self.obj.SetInputValue(0, code)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, listWatchData)
        self.obj.Subscribe()


# CpMarketWatchS:
class CpMarketWatchS(CpPublish):
    def __init__(self):
        super().__init__('CpSysDib.CpMarketWatchS')


# CpRpMarketWatch : 특징주 포착 통신
class CpRpMarketWatch:
    def __init__(self):
        self.cpMarketWatch = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
        self.cpMarketWatchS = CpMarketWatchS()

    def Request(self, code, listWatchData: OrderedDict):
        self.cpMarketWatchS.Unsubscribe()

        self.cpMarketWatch.SetInputValue(0, code)
        # 1: 종목 뉴스 2: 공시정보 10: 외국계 창구첫매수, 11:첫매도 # 12 외국인 순매수 13 순매도
        self.cpMarketWatch.SetInputValue(1, '1,2,10,11,12,13')
        self.cpMarketWatch.SetInputValue(2, 0)  # 시작 시간: 0 처음부터

        self.cpMarketWatch.BlockRequest()

        for i in range(self.cpMarketWatch.GetHeaderValue(2)):
            code = self.cpMarketWatch.GetDataValue(1, i)
            indicator = self.cpMarketWatch.GetDataValue(3, i)
            item = {
                'code': code,
                'indicator': indicator
            }
            if indicator in indicators:
                item['remark'] = indicators[indicator]
            listWatchData[code] = item

        self.cpMarketWatchS.Subscribe(code, listWatchData)
