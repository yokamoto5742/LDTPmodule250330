import subprocess
import os
import shutil
from version_manager import update_version, update_version_py


def build_executable():
    new_version = update_version()
    update_version_py(new_version)

    subprocess.run([
        "pyinstaller",
        "--name=LDTPapp",
        "--windowed",
        "--icon=assets/LDPTapp_icon.ico",
        "main.py"
    ])

    dist_path = os.path.join('dist', 'LDTPapp')

    if not os.path.exists(os.path.join(dist_path, 'config.ini')):
        shutil.copy('config.ini', dist_path)

    print(f"Executable built successfully. Version: {new_version}")


if __name__ == "__main__":
    build_executable()
