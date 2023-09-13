import subprocess
import time

ppt_file1 = ""
ppt_file2 = ""

launch_time1 = "", ""
launch_time2 = ""

current_time =time.strftime("%H:%M:%S")

if current_time == launch_time1:
    print(f"lzncement de {ppt_file1} à {launch_time1}.")
    time.sleep()