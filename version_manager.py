import os
import re
from datetime import datetime

VERSION_FILE = "version.txt"

def get_current_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "r") as f:
            return f.read().strip()
    return "0.0.0"

def increment_version(version):
    major, minor, patch = map(int, version.split("."))
    return f"{major}.{minor}.{patch + 1}"

def update_version():
    current_version = get_current_version()
    new_version = increment_version(current_version)

    with open(VERSION_FILE, "w") as f:
        f.write(new_version)

    return new_version

def update_version_py(new_version):
    with open('version.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # バージョン情報を更新
    content = re.sub(r'VERSION = "[0-9.]+"', f'VERSION = "{new_version}"', content)

    # 最終更新日を更新
    today = datetime.now().strftime("%Y/%m/%d")
    content = re.sub(r'LAST_UPDATED = "[0-9/]+"', f'LAST_UPDATED = "{today}"', content)

    with open('version.py', 'w', encoding='utf-8') as f:
        f.write(content)
