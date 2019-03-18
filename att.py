import cv2
import math
import os
import shutil
from PIL import Image
from pyzbar.pyzbar import decode
import xlwt 
from xlwt import Workbook

cap = cv2.VideoCapture(0)

if not cap.isOpened():
    raise IOError("Cannot open webcam")


path = "/Users/nikhilanu/Desktop/projects/opencv/data"

try:  
    shutil.rmtree(path)
except OSError:  
    print ("Deletion of the directory %s failed" % path)
else:  
    print ("Successfully deleted the directory %s" % path)

try:  
    os.makedirs(path)
except OSError:  
    print ("Creation of the directory %s failed" % path)
else:  
    print ("Successfully created the directory %s" % path)

currentframe = 0
l=[]  
j=0
while True:
    ret, frame = cap.read()
    frame = cv2.resize(frame, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
    cv2.imshow('Input', frame)

    if ret:
        name = './data/frame' + str(currentframe) + '.jpg'
        #print ('Creating...' + name)
        cv2.imwrite(name, frame)
        decoded = decode(Image.open('./data/frame' + str(currentframe) + '.jpg'))
        if decoded!=[]:
            data=decoded[0].data
            data=data.decode('utf-8')
            if data in l:
                pass
            else:
                l.append(data)
                j=j+1
                print(data)
        currentframe += 1
    else: 
        break
    
    c = cv2.waitKey(1)
    if c == 27:
        break


wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
  
for k in range (0,j):
    words = l[k].split()
    x=len(words)
    for z in range (0,x):
        sheet1.write(k, z, words[z]) 
   

wb.save('/Users/nikhilanu/Desktop/xlwt example.xls')

cap.release()
cv2.destroyAllWindows()
