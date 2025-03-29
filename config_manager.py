import configparser
import os
import sys
from typing import Any


def get_config_path():
    if getattr(sys, 'frozen', False):
        # PyInstallerでビルドされた実行ファイルの場合
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, 'config.ini')

CONFIG_PATH = get_config_path()



def load_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            config.read_file(f)
    except FileNotFoundError:
        print(f"設定ファイルが見つかりません: {CONFIG_PATH}")
        raise
    except PermissionError:
        print(f"設定ファイルを読み取る権限がありません: {CONFIG_PATH}")
        raise
    except configparser.Error as e:
        print(f"設定ファイルの解析中にエラーが発生しました: {e}")
        raise
    return config


def save_config(config: configparser.ConfigParser):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
    except PermissionError:
        print(f"設定ファイルを書き込む権限がありません: {CONFIG_PATH}")
        raise
    except IOError as e:
        print(f"設定ファイルの保存中にエラーが発生しました: {e}")
        raise
