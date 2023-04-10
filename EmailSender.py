from pdf_mail import sendpdf
import os

#Details of SenderMailID, ReceiverMailID & DocInformation 
k = sendpdf("kabaliproctorsystem@yahoo.com",    # Sender mail ID
            "modthoufeeq@gmail.com",            # receiver mail ID
            "TVSsystem@007",                    # Sender mail ID password
            "Proctor Report",                   # Title of Email
            "This the detailed report for the interview held now.\n Thank you for using KABALI proctoring system.\nYour Feedback is our priority.",
            "Interview Report",                 # File name
            "D:/KABALI/KAB/Backend/IO Devices") # Absolute path(wkhtmltopdf path)

k.email_send()
