from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier(r"C:\Users\archa\OneDrive\Documents\face_recognition_project-main\face_recognition_project-main\data\haarcascade_frontalface_default.xml")

# Load the LABELS and FACES data
with open(r"C:\Users\archa\OneDrive\Documents\face_recognition_project-main\face_recognition_project-main\data\names.pkl", 'rb') as w:
    LABELS = pickle.load(w)
with open(r"C:\Users\archa\OneDrive\Documents\face_recognition_project-main\face_recognition_project-main\data\faces.pkl", 'rb') as f:
    FACES = pickle.load(f)

# Ensure FACES and LABELS have the same length and contain only "kalpana"
LABELS = ["kalpana"] * len(FACES)

print('Shape of Faces matrix --> ', FACES.shape)
print('Number of Labels --> ', len(LABELS))

# Train the KNN classifier
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground = cv2.imread(r"C:\Users\archa\OneDrive\Documents\face_recognition_project-main\face_recognition_project-main\background.png")

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        
        # Capture the timestamp
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        attendance = ["kalpana", str(timestamp)]

        # Check if attendance file exists
        exist = os.path.isfile(f"Attendance/Attendance_{date}.csv")
        
        # Draw bounding box and label on detected face
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, "kalpana", (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

    # Overlay the video frame onto the background image
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)
    
    # Save attendance when 'o' is pressed
    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)

    # Exit the loop when 'q' is pressed
    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
