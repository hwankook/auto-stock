import xml.etree.ElementTree as ET
from datetime import datetime

import requests

from config import config


def get_request_query(params):
    import urllib.parse as urlparse
    params = urlparse.urlencode(params)
    query = 'http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo?' + params + '&' \
            + 'serviceKey' + '=' + config.service_key
    return query


def is_holiday():
    date = datetime.now()

    weekday = date.weekday()
    if weekday == 5 or weekday == 6:
        return True

    solYear = date.year
    solMonth = f'{date.month:02d}'
    params = {'solYear': solYear, 'solMonth': solMonth}

    request_query = get_request_query(params)
    # print('request:', request_query)

    response = requests.get(url=request_query)
    # print('status_code:', response.status_code)

    if response.ok:
        result = response.text

        root = ET.fromstring(result)

        today = str(solYear) + str(solMonth) + f'{date.day:02d}'
        # print(today)
        items = root.find('./body/items')
        for item in items:
            isHoliday = item.find('isHoliday').text
            locdate = item.find('locdate').text
            # print(locdate, isHoliday)
            if today == locdate and isHoliday == 'Y':
                print('holiday')
                return True

    return False


if __name__ == '__main__':
    print('Is holiday?', is_holiday())
