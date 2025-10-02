import os
import sys
import subprocess
import time
import psutil

def main():
    """守护进程主函数"""
    # 获取当前脚本路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(current_dir, "main.pyw")
    while True:
        process = subprocess.Popen(["pythonw", main_script])
        process.wait()
        time.sleep(3)
if __name__ == "__main__":
    main()