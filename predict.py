import csv
import os
import sys
import traceback

import pandas as pd
from fbprophet import Prophet


def predict_price():
    """Prophet으로 당일 종가 가격 예측"""
    path = './ohlc'

    for file_path in os.listdir(path):
        ohlc = pd.read_csv(path + '/' + file_path)

        if len(ohlc) <= 30:
            continue

        ohlc['ds'] = pd.to_datetime(ohlc.iloc[:, 0].astype(str), format='%Y-%m-%d')
        ohlc['y'] = ohlc.iloc[:, 4]
        data = ohlc[['ds', 'y']]

        model = Prophet()
        model.fit(data)

        future = model.make_future_dataframe(periods=1, freq='D')

        forecast = model.predict(future)

        code = file_path.split('-')[0]
        price = forecast['yhat'].values[-1]

        try:
            os.makedirs('./predict', exist_ok=True)

            with open('./predict/' + code + '.csv', 'a', encoding="utf-8", newline='\n') as f:
                csv_writer = csv.writer(f)
                row = [code, price]
                csv_writer.writerow(row)
        except Exception:
            traceback.print_exc(file=sys.stdout)


if __name__ == '__main__':
    predict_price()
    # schedule.every().hour.do(lambda: predict_price("KRW-BTC"))
