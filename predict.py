import csv
import os
import sys
import time
import traceback
from datetime import datetime

import pandas as pd
import schedule
from fbprophet import Prophet


def predict_price():
    """Prophet으로 당일 종가 가격 예측"""
    path = './curr'

    for file_path in os.listdir(path):
        df = pd.read_csv(path + '/' + file_path)

        model = Prophet()
        model.fit(df)

        future = model.make_future_dataframe(periods=1, freq='H')

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
    """자동 매도, 매수, 종료한다."""
    t_now = datetime.now()
    t_start = t_now.replace(hour=8, minute=30, second=0, microsecond=0)
    t_exit = t_now.replace(hour=16, minute=00, second=0, microsecond=0)

    if t_start < t_now < t_exit:
        predict_price()

        schedule.every(10).minutes.do(lambda: predict_price())

        while True:
            schedule.run_pending()
            time.sleep(1)
