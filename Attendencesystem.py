import face_recognition
import cv2
import numpy as np
import datetime
import openpyxl
import json
import os

from twilio.rest import Client


with open('config.json') as config_file:
    config = json.load(config_file)


account_sid = "AC5fa96edb42c487a9323b86bd3c006b42"
auth_token = "8951afed2cada840d9ba626755dc73a1"
twilio_phone_number = '+1 334 219 9854'

client = Client(account_sid, auth_token)


video_capture = cv2.VideoCapture(0)


known_face_encodings = []
known_face_names = []
parent_numbers = {}

for person in config['people']:
    image = face_recognition.load_image_file(person['image_path'])
    face_encoding = face_recognition.face_encodings(image)[0]
    known_face_encodings.append(face_encoding)
    known_face_names.append(person['name'])
    parent_numbers[person['name']] = person['parent_number']


face_locations = []
face_encodings = []
face_names = []
face_confidences = [] 
present_names = set()  
absent_names = set(known_face_names) 
on_duty_names = set(config['on_duty_students'])  


workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet["A1"] = "Name"
worksheet["B1"] = "Timestamp"
worksheet["C1"] = "Status" 
row = 2

try:
    while True:
       
        ret, frame = video_capture.read()

       
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

       
        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

       
        rgb_small_frame = small_frame[:, :, ::-1]

        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        face_names = []
        face_confidences = []  

        for face_encoding in face_encodings:
           
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"
            confidence = 0  

           
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            if matches[best_match_index]:
                name = known_face_names[best_match_index]
                confidence = (1 - face_distances[best_match_index]) * 100
                if name in absent_names and name not in present_names and confidence > 50:
                    present_names.add(name) 
                    worksheet[f"A{row}"] = name
                    if name in on_duty_names:
                        worksheet[f"B{row}"] = "OD"  
                        worksheet[f"C{row}"] = "Present"  
                    else:
                        worksheet[f"B{row}"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        worksheet[f"C{row}"] = "Present"  
                    row += 1

            face_names.append(name)
            face_confidences.append(confidence)

       
        for name in on_duty_names:
            if name not in face_names and name not in present_names:
                present_names.add(name) 
                worksheet[f"A{row}"] = name
                worksheet[f"B{row}"] = "OD"  
                worksheet[f"C{row}"] = "Present"  
                row += 1

      
        for (top, right, bottom, left), name, confidence in zip(face_locations, face_names, face_confidences):
          
            top *= 4
            right *= 4
            bottom *= 4
            left *= 4

           
            if confidence > 50:
                color = (0, 255, 0)  
                attendance_status = "Present"
            else:
                color = (0, 0, 255) 
                attendance_status = "Absent"

            cv2.rectangle(frame, (left, top), (right, bottom), color, 2)

           
            cv2.rectangle(frame, (left, bottom - 35), (right, bottom), color, cv2.FILLED)
            font = cv2.FONT_HERSHEY_DUPLEX
            cv2.putText(frame, f"{name}: {confidence:.2f}%", (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

       
        cv2.imshow('Video', frame)
finally:
   
    absent_names_list = list(absent_names - present_names)  
    for name in absent_names_list:
        worksheet[f"A{row}"] = name
        worksheet[f"B{row}"] = "Absent"
        worksheet[f"C{row}"] = "Absent"  
        row += 1
    workbook.save("D:/go/attendance.xlsx")
    print("Attendance Successful")


    for name in absent_names_list:
        parent_number = parent_numbers.get(name)
        if parent_number:
            message = f"Your child, {name}, is absent today."
            client.messages.create(body=message, from_=twilio_phone_number, to=parent_number)
            print(f"SMS sent to {parent_number} for {name}")

 
    video_capture.release()
    cv2.destroyAllWindows()



import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email(sender_email, sender_password, receiver_email, subject, body, attachment_path):
   
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

   
    with open(attachment_path, "rb") as attachment:
       
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

   
    encoders.encode_base64(part)

    
    part.add_header("Content-Disposition", "attachment", filename=attachment_path)

   
    message.attach(part)
    text = message.as_string()

   
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, text)


sender_email = "manikumar0459@gmail.com"
sender_password = "ftie wnlb hfbl xeyb"
receiver_email = "veldurthivenkatamani@gmail.com"
subject = "Attendance Report"
body = "ATTENDANCE REPORT"
attachment_path = os.path.join("attendance.xlsx")
send_email(sender_email, sender_password, receiver_email, subject, body, attachment_path)
print("Email sent successfully")
