import configparser
import os

__all__ = ['config']

parser = configparser.ConfigParser(interpolation=None)
parser.read(os.path.dirname(os.path.abspath(__file__)) + '/config.ini')


class Config(object):
    __instance = None

    def __new__(cls):
        if Config.__instance is None:
            Config.__instance = object.__new__(cls)

        Config.__instance.id = parser['DEFAULT']['id']
        Config.__instance.pwd = parser['DEFAULT']['pwd']
        Config.__instance.pwdcert = parser['DEFAULT']['pwdcert']
        Config.__instance.token = parser['DEFAULT']['token']
        Config.__instance.service_key = parser['DEFAULT']['ServiceKey']
        Config.__instance.code_limit = int(parser['DEFAULT']['code_limit'])
        Config.__instance.target_buy_count = int(parser['DEFAULT']['target_buy_count'])
        Config.__instance.buy_amount = int(parser['DEFAULT']['buy_amount'])
        Config.__instance.profit_rate = float(parser['DEFAULT']['profit_rate'])
        Config.__instance.loss_rate = float(parser['DEFAULT']['loss_rate'])
        Config.__instance.K = float(parser['DEFAULT']['K'])

        return Config.__instance


config = Config()
