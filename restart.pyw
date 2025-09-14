import time
import os
import subprocess
subprocess.run(["taskkill","/f","/im","contest_helper.exe"])
time.sleep(1)
current_script = os.path.join(os.path.dirname(__file__), "main.pyw")
subprocess.Popen(["pythonw", current_script])