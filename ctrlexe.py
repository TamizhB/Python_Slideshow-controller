# -*- coding: utf-8 -*-
"""
Created on Sun Feb  4 22:29:27 2018

@author: Tamil SB
"""
import numpy as np
import cv2
import win32com.client
import time
def run(filename):
    x = win32com.client.Dispatch("PowerPoint.Application")
    x.Presentations.Open(FileName=filename)
    x.ActivePresentation.SlideShowSettings.Run()
    slide_count=x.ActivePresentation.Slides.count
    slide_no=1
    cv2.namedWindow('mask',cv2.WINDOW_NORMAL)
    cv2.resizeWindow('mask', 200,200)
    cv2.namedWindow('fin',cv2.WINDOW_NORMAL)
    cv2.resizeWindow('fin', 200,200)
    class Foreground_Extractor:
        def __init__(self,alpha,firstFrame):
            self.alpha=alpha
            self.background=firstFrame
        def getForeground(self,frame):
            self.background=frame * self.alpha + self.background * (1 - self.alpha)
            return cv2.absdiff(self.background.astype(np.uint8),frame)
    cam=cv2.VideoCapture(0)
    time.sleep(2)
    t=1
    upd=1
    center=(0,0)
    list_=[]
    def denoise(frame):
        frame = cv2.medianBlur(frame,5)
        frame = cv2.GaussianBlur(frame,(5,5),0)
        return frame
    ret,frame=cam.read()
    if ret is True:
        extractor=Foreground_Extractor(0.01,frame)
        run=True
    else:
        run=False
    while(run):
        ret,frame=cam.read()
        #cv2.imshow('input',frame)
        foreground=extractor.getForeground(frame)
        imgray=cv2.cvtColor(foreground,cv2.COLOR_BGR2GRAY)
        ret, mask = cv2.threshold(imgray,20, 255, cv2.THRESH_BINARY)
        res = cv2.bitwise_and(frame,frame,mask=mask)
        hsv=cv2.cvtColor(res,cv2.COLOR_BGR2HSV)
        lower_limit = np.array([0,50,60])
        upper_limit = np.array([60,175,178])
        mask1 = cv2.inRange(hsv, lower_limit, upper_limit)
        blur = cv2.GaussianBlur(mask1,(5,5),0)
        final = cv2.medianBlur(blur, 5)
        ret3,thresh = cv2.threshold(final,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)
        image,contours, hierarchy = cv2.findContours(thresh,cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
        c=0
        for i in range(0,len(contours)):
            l=len(contours[i])
            if l>c:
                c=l
                d=i

        #cv2.drawContours(frame,contours[d],-1,(0,255,0),2)
        area = cv2.contourArea(contours[d])
        if area>=3000:
            M = cv2.moments(contours[d])
            cX = int(M['m10'] /M['m00'])
            cY = int(M['m01'] /M['m00'])
            center=(cX,cY)
            cv2.drawContours(frame,contours[d],-1,(0,255,0),2)
            cv2.circle(frame, center, 5, [0,0,255], -1)
        else:
            center=(0,0)
        if center !=(0,0):
            t+=1
            if t%6==0:
                    #print(center)
                list_.append(center[0])
                upd+=1
                if upd==4:
                    upd=1
                    t=1
                    center=(0,0)
            
                    count_=list(np.diff(list_) > 0)
                    if np.all(abs(np.diff(list_))<=5):
                        break
                    list_.clear()
                    if count_.count(True)-count_.count(False)>0:
                        if(slide_no>1):
                        #print("Left")
                            slide_no-=1
                            x.SlideShowWindows(1).View.Previous()
                    else:
                        if(slide_no<=slide_count):
                        #print("Right")
                            slide_no+=1
                            x.SlideShowWindows(1).View.Next()
        else:
            upd=1
            list_.clear()
        cv2.imshow('mask',thresh)
        cv2.imshow('fin',frame)
        key = cv2.waitKey(10) & 0xFF
    
        if key == 27:
            break

#x.SlideShowWindows(1).View.Exit()
    x.Quit()
    cam.release()
    cv2.destroyAllWindows()
