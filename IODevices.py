import cv2
import pyaudio
import pygame
import pyautogui
import platform
import wmi
import subprocess
import bluetooth 

# To establish connection with Mongo Db
import pymongo
from bson.binary import Binary

# Establish a connection to the MongoDB server
client = pymongo.MongoClient("mongodb://localhost:27017/")

# Select the database and collection to work with
db = client["KABALIDB"]

#drop the collection if there is already one existing
if "IODevices" in db.list_collection_names():
    # Get the "customer" collection object
    IODevices_collection = db["IODevices"]
    # Drop the "customer" collection
    IODevices_collection.drop()

collection = db["IODevices"]

# Get the number of monitors
if platform.system() == "Windows":
    import win32api

    monitor_count = len(win32api.EnumDisplayMonitors())
    Monitor = {"name": "Monitor Connected", "data": monitor_count}
elif platform.system() == "Darwin":
    command = "system_profiler SPDisplaysDataType | grep Resolution | wc -l"
    output = subprocess.check_output(command, shell=True)
    monitor_count = int(output)
    Monitor = {"name": "Monitor Connected", "data": monitor_count}
elif platform.system() == "Linux":
    command = "xrandr --listactivemonitors | grep -c Monitors"
    output = subprocess.check_output(command, shell=True)
    monitor_count = int(output)
    Monitor = {"name": "Monitor Connected", "data": monitor_count}
else:
    Monitor = {"name": "Monitor Connected", "data": "Unsupported Monitor Connected"}
    monitor_count = 0
# Insert the data into the document
collection.insert_one(Monitor)
# Successfully inserted the Monitor count

# Webcam detection
webcam_count = 0
while True:
    cap = cv2.VideoCapture(webcam_count)
    if not cap.read()[0]:
        break
    webcam_count += 1
Webcam = {"name": "Webcam Connected", "data": webcam_count}
# Insert the data into the document
collection.insert_one(Webcam)
# Successfully inserted the Webcam count

# Mouse detection
try:
    pyautogui.size()
    mouse_count = 0
    if platform.system() == "Windows":
        c = wmi.WMI()
        for device in c.Win32_PnPEntity():
            if 'mouse' in str(device.Caption).lower():
                mouse_count += 1
    elif platform.system() == "Darwin":
        command = "system_profiler SPUSBDataType | grep 'Manufacturer:\\|Product:' | awk '{print $NF}' | sed 's/USB//g' | sed 'N;s/\\n/ /'"
        output = subprocess.check_output(command, shell=True)
        mouse_count = output.count("Mouse")
    elif platform.system() == "Linux":
        command = "xinput list | grep -c 'mouse\\|Mouse'"
        output = subprocess.check_output(command, shell=True)
        mouse_count = int(output)
    Mouse = {"name": "Mouse Connected", "data": mouse_count}
    # Insert the data into the document
    collection.insert_one(Mouse)
    # Successfully inserted the Mouse count
except:
    Mouse = {"name": "Mouse Connected", "data": "No Mouse Detected"}

# Keyboard detection
try:
    pygame.init()
    pygame.display.set_mode((100, 100))
    keyboard_count = 1
    while True:
        pygame.event.pump()
        if not pygame.key.get_focused():
            break
        keyboard_count += 1
    Keyboard_data = {"name": "Keyboard Connected", "data": keyboard_count}
    # Insert the data into the document
    collection.insert_one(Keyboard_data)
    # Successfully inserted the Keyboard count
except:
    Keyboard_data = {"name": "Keyboard Connected", "data": "No Keyboard detected"}

# Speakers and microphone detection
p = pyaudio.PyAudio()
info = p.get_host_api_info_by_index(0)
num_devices = info.get('deviceCount')
speaker_count = 0
microphone_count = 0
for i in range(num_devices):
    device_info = p.get_device_info_by_index(i)
    max_output_channels = device_info.get('maxOutputChannels')
    max_input_channels = device_info.get('maxInputChannels')
    if max_output_channels > 0 and max_input_channels == 0:
        if i == p.get_default_output_device_info()['index']:
            speaker_count += 1
    elif max_output_channels == 0 and max_input_channels > 0:
        if 'microphone' in device_info.get('name').lower():
            microphone_count += 1
Speakers = {"name": "Speakers Connected", "data": speaker_count}
Microphone = {"name": "Microphone Connected", "data": microphone_count}
# Insert the data into the document
collection.insert_one(Speakers)
collection.insert_one(Microphone)
# Successfully inserted the Speakers and Microphone count

# Get a list of connected devices
# connected_devices = []
# bluetoothCount = 0
# for addr in bluetooth.discover_devices():
#     services = bluetooth.find_service(address=addr)
#     for service in services:
#         if service['host'] == addr:
#             # Check if the device is already in the list
#             if addr not in [dev[0] for dev in connected_devices]:
#                 connected_devices.append((addr, bluetooth.lookup_name(addr)))
#                 bluetoothCount+=1

# # Print the names and addresses of the connected devices
# if len(connected_devices) > 0:
#     for addr, name in connected_devices:
#         Bluetooth_Count = {"name": "Bluetooth Devices Connected", "data": bluetoothCount}
#         Bluetooth = {"name": "Bluetooth Devices Information", "DeviceName": name, "Address": addr}
#         # Insert the data into the document
#         collection.insert_one(Bluetooth_Count)
#         collection.insert_one(Bluetooth)
#         # Successfully inserted the Bluetooth Devices count
# else:
#     Bluetooth = {"name": "Bluetooth Devices Connected", "data": 0}
#     collection.insert_one(Bluetooth)