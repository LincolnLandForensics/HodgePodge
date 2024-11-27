import os
import shutil
from datetime import datetime
import numpy as np
# import pandas as pd
import xlsxwriter
from imutils.object_detection import non_max_suppression
import cv2

'''
This is a clone of
https://github.com/ndaidong/people-detecting.git
I am trying to add spreadsheet output for counting people

'''




version = '1.0.1'

hog = cv2.HOGDescriptor()
hog.setSVMDetector(cv2.HOGDescriptor_getDefaultPeopleDetector())

font = cv2.FONT_HERSHEY_SIMPLEX

dest_folder = "nobody"

# Create a new Excel file
workbook = xlsxwriter.Workbook('videos.xlsx')
worksheet = workbook.add_worksheet()

# Set column headers
worksheet.write(0, 0, 'video')
worksheet.write(0, 1, 'People (approximate)')
worksheet.write(0, 2, 'FileSize (bytes)')
# worksheet.write(0, 3, 'creation_time')
# worksheet.write(0, 4, 'access_time')
# worksheet.write(0, 5, 'modified_time')

# freeze cells at 1,1
worksheet.freeze_panes(1, 1)
# set column widths
worksheet.set_column(0, 0, 25)
worksheet.set_column(1, 1, 20)
worksheet.set_column(2, 2, 16)
worksheet.set_column(3, 3, 20)
worksheet.set_column(4, 4, 20)
worksheet.set_column(5, 5, 20)

def move_file(file_path, dest_folder):
    if os.path.isfile(file_path):
        # file_size = os.path.getsize(file_path)
        if 1==1:
        # if file_size < 20 * 1024: # 20 KB
            shutil.move(file_path, dest_folder)
            print(f"File {file_path} moved to {dest_folder} because nobody was in it.")
        else:
            print(f"File {file_path} is too big to move")
    else:
        print(f"{file_path} is not a valid file")



# Iterate through all vidoes in the "videos" folder
row = 1
for file_name in os.listdir('videos'):
    (video, file_size) = ('', '')
    video = os.path.join('videos', file_name)
    # if file_name.endswith('.avi'):
        # video = os.path.join('videos', file_name)

    file_size = os.path.getsize(video)

    # get creation, access, modified times of the .plist
    file_info = os.stat(video)

    # utc date
    creation_time = os.path.getctime(video)
    creation_time_utc = datetime.utcfromtimestamp(creation_time)

    modified_time = os.path.getmtime(video)
    modified_time_utc = datetime.utcfromtimestamp(modified_time)



    # video = 'vtest2.avi'
    cap = cv2.VideoCapture(video)

    w = cap.get(3)
    h = cap.get(4)
    mx = int(w - 400)
    my = int(h - 24)

    results = []

    people_detected = 0
    max_people_detected = 0
        
    while(cap.isOpened()):
        ret, frame = cap.read()

        if ret is False:
            break
        
        k = cv2.waitKey(30) 
        if k == 27:
            break

        rects, weights = hog.detectMultiScale(
            frame,
            winStride=(4, 4), 
            padding=(8, 8), 
            scale=1.05
        )
        for (x, y, w, h) in rects:
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)

        rects = np.array([[x, y, x + w, y + h] for (x, y, w, h) in rects])
        pick = non_max_suppression(rects, probs=None, overlapThresh=0.65)

        people_detected = 0
        
        for (xA, yA, xB, yB) in pick:
            people_detected += 1
            cv2.rectangle(frame, (xA, yA), (xB, yB), (0, 255, 0), 2)

        if people_detected > max_people_detected:
            max_people_detected = people_detected

        text = 'People detected: ' + str(people_detected)
        cv2.putText(
            frame, 
            text, 
            (mx, my), 
            font, 
            1, 
            (255, 255, 255), 
            1, 
            cv2.LINE_AA
        )
        cv2.imshow('Frame', frame)

    results.append({'Video': video, 'People Detected': max_people_detected, 'FileSize': file_size})
    worksheet.write(row, 0, file_name)
    worksheet.write(row, 1, max_people_detected)
    worksheet.write(row, 2, file_size)
    # worksheet.write(row, 3, creation_time_utc)
    # worksheet.write(row, 4, access_time)
    # worksheet.write(row, 5, modified_time_utc)            
    row += 1
    cap.release()
    cv2.destroyAllWindows()

    # file_path = "path/to/file.txt"

    if max_people_detected == 0:
        move_file(video, dest_folder)


# Write results to an excel file
# df = pd.DataFrame(results)
# df.to_excel('output.xlsx', index=False)

# Save and close the Excel file
workbook.close()
