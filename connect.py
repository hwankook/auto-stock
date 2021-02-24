import os
import sys
import time
import traceback

from pywinauto import application

from config import config


def connect():
    try:
        # os.system('taskkill /IM coStarter* /F /T')
        # os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        time.sleep(5)

        app = application.Application()
        app.start(f'C:\CREON\STARTER\coStarter.exe /prj:cp '
                  f'/id:{config.id} /pwd:{config.pwd} /pwdcert:{config.pwdcert} /autostart')
        time.sleep(180)
    except OSError as e:
        traceback.print_exc(file=sys.stdout)
        print('`connect -> exception! ' + str(e) + '`')


if __name__ == '__main__':
    connect()
