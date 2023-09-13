import subprocess
from subprocess import TimeoutExpired
import time

# Define the paths to your PPT files
ppt_file1 = "Présentation1.pptx"
ppt_file2 = "Présentation2.pptx"

# Define the specific times when you want to launch the PPT files (24-hour format)
launch_time1 = "09/12/2023 20:06:40" #, "18:17:00"  # Replace with your desired time
launch_time2 = "09/12/2023 20:07:00"  # Replace with your desired time
launch_time3 = "09/12/2023 20:08:00"

# Get the current time
current_time = time.strftime("%m/%d/%Y %H:%M:%S")
print(f"current time {current_time}")

# Define a function to calculate the time difference in seconds
def time_difference(target_time):
    target = time.strptime(target_time, "%m/%d/%Y %H:%M:%S")
    current = time.strptime(current_time, "%m/%d/%Y %H:%M:%S")
    return (time.mktime(target) - time.mktime(current))

# Calculate the time difference in seconds until the launch
time_until_launch1 = time_difference(launch_time1)
time_until_launch2 = time_difference(launch_time2)
time_until_launch3 = time_difference(launch_time3)

# Sleep until the first launch time
if time_until_launch1 > 0:
    print(f"Waiting {time_until_launch1} seconds until the first launch...")
    time.sleep(time_until_launch1)

# Launch the first PPT file
proc1 = subprocess.Popen(["start", "powerpnt", ppt_file1], shell=True)

# Sleep until the second launch time
if time_until_launch2 > 0:
    print(f"Waiting {time_until_launch2} seconds until the second launch...")
    time.sleep(time_until_launch2)

# Close file
try:
    outs, errs = proc1.communicate(timeout=15)
except TimeoutExpired:
    proc1.kill()
    outs, errs = proc1.communicate()

# Launch the second PPT file
proc2 = subprocess.Popen(["start", "powerpnt", ppt_file2], shell=True)

if time_until_launch3 > 0:
    print(f"Waiting {time_until_launch3} seconds until the first launch...")
    time.sleep(time_until_launch1)

try:
    outs, errs = proc2.communicate(timeout=15)
except TimeoutExpired:
    proc2.kill()
    outs, errs = proc2.communicate()

# Launch the first PPT file
subprocess.Popen(["start", "powerpnt", ppt_file1], shell=True)
