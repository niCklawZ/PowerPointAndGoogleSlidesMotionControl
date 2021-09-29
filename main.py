import tkinter.filedialog

import win32com.client
import cv2
from cvzone.HandTrackingModule import HandDetector
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = tkinter.filedialog.askopenfilename(filetypes=[("PowerPoint files", ".pptx")]).replace("/", "\\")
print(file_path)

app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
app.DisplayAlerts = True
presentation = app.Presentations.Open(FileName=''+file_path)
presentation.SlideShowSettings.ShowPresenterView = False



cap = cv2.VideoCapture(0)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)

detector = HandDetector(detectionCon=0.8, maxHands=2)
action = False
presensationMode = False
while True:
    success, img = cap.read()
    hands, img = detector.findHands(img)
    cv2.imshow("Image", img)
    cv2.waitKey(1)

    if len(hands) == 1:
        if detector.fingersUp(hands[0]) == [1, 1, 1, 1, 1]:
            if not action:
                presensationMode = not presensationMode
                action = True
                if presensationMode:
                    presentation.SlideShowSettings.Run()
                elif not presensationMode:
                    presentation.SlideShowWindow.View.Exit()
        elif detector.fingersUp(hands[0]) == [0, 1, 0, 0, 0]:
            if not action:
                print("vorwärts")
                if presensationMode:
                    presentation.SlideShowWindow.View.Next()
                action = True
        elif detector.fingersUp(hands[0]) == [0, 1, 1, 0, 0]:
            if not action:
                print("rückwärts")
                if presensationMode:
                    presentation.SlideShowWindow.View.Previous()
                action = True
        elif detector.fingersUp(hands[0]) == [1, 0, 0, 0, 1]:
            if not action:
                action = True
                app.Quit()
                exit()
        else:
            action = False
    elif len(hands) == 2:
        if detector.fingersUp(hands[0]) == [1, 1, 1, 1, 1] or detector.fingersUp(hands[1]) == [1, 1, 1, 1, 1]:
            if not action:
                presensationMode = not presensationMode
                action = True
                if presensationMode:
                    presentation.SlideShowSettings.Run()
                elif not presensationMode:
                    presentation.SlideShowWindow.View.Exit()
        elif detector.fingersUp(hands[0]) == [0, 1, 0, 0, 0] or detector.fingersUp(hands[1]) == [0, 1, 0, 0, 0]:
            if not action:
                print("vorwärts")
                if presensationMode:
                    presentation.SlideShowWindow.View.Next()
                action = True
        elif detector.fingersUp(hands[0]) == [0, 1, 1, 0, 0] or detector.fingersUp(hands[1]) == [0, 1, 1, 0, 0]:
            if not action:
                print("rückwärts")
                if presensationMode:
                    presentation.SlideShowWindow.View.Previous()
                action = True
        elif detector.fingersUp(hands[0]) == [1, 0, 0, 0, 1] or detector.fingersUp(hands[1]) == [1, 0, 0, 0, 1]:
            if not action:
                action = True
                app.Quit()
                exit()
        else:
            action = False
    else:
        action = False






























'''
while True:
    if keyboard.read_key() == "a":
        presentation.SlideShowWindow.View.Previous()
        time.sleep(0.01)
    elif keyboard.read_key() == "s":
        presentation.SlideShowWindow.View.Next()
        time.sleep(0.01)



time.sleep(1)
presentation.SlideShowWindow.View.Next()
time.sleep(1)
presentation.SlideShowWindow.View.Next()
time.sleep(1)
presentation.SlideShowWindow.View.Previous()
time.sleep(1)

presentation.SlideShowWindow.View.Exit()

app.Quit()
'''
