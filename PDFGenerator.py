import jinja2
import pdfkit
import pymongo
from datetime import datetime
import os

# To establish connection with Mongo Db
client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["KABALIDB"]
IOcollection = db["IODevices"]
# retrieve a document from the collection
query1 = {"name": "Monitor Connected"}
document1 = IOcollection.find_one(query1)
query2 = {"name": "Keyboard Connected"}
document2 = IOcollection.find_one(query2)
query3 = {"name": "Mouse Connected"}
document3 = IOcollection.find_one(query3)
query4 = {"name": "Webcam Connected"}
document4 = IOcollection.find_one(query4)
query5 = {"name": "Speakers Connected"}
document5 = IOcollection.find_one(query5)
query6 = {"name": "Microphone Connected"}
document6 = IOcollection.find_one(query6)
# query7 = {"name": "Bluetooth Devices Connected"}
# document7 = IOcollection.find_one(query7)

# get the value of a particular field
monitorCount = int(document1["data"])
keyboardCount = int(document2["data"])
mouseCount = int(document3["data"])
webCameraCount = int(document4["data"])
speakerCount = int(document5["data"])
microphoneCount = int(document6["data"])
# bluetoothDevicesCount = int(document7["data"])

#setting the context in dictionary format
context = {'monitorCount': monitorCount, 'keyboardCount': keyboardCount, 'mouseCount': mouseCount, 'webCameraCount': webCameraCount,
           'speakerCount': speakerCount, 'microphoneCount': microphoneCount}

#checking if there is any report of previous session is present. If YES then removing it from the disk
if os.path.exists("./Interview Report.pdf"):
    os.remove("Interview Report.pdf")

#placing the template of our report as html
template_loader = jinja2.FileSystemLoader('./')
template_env = jinja2.Environment(loader=template_loader)

#rendering the values in html file to the value present in context list
html_template = 'report.html'
template = template_env.get_template(html_template)
output_text = template.render(context)

#generating a new pdf with the value
config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')
output_pdf = 'Interview Report.pdf'
pdfkit.from_string(output_text, output_pdf, configuration=config)