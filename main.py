import time
import tkinter.filedialog
import win32com.client
import cv2
from cvzone.HandTrackingModule import HandDetector
import tkinter as tk
import webbrowser
import keyboard

cam_id = int(input("Enter camera id (0..n): "))

presentationType = input("Type in 'pp' for PowerPoint or 'gs' for Google Slides: ")
presentationType = presentationType.lower()

if presentationType == "pp" or presentationType == "gs":
    if presentationType == "pp":
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        file_path = tkinter.filedialog.askopenfilename(filetypes=[("PowerPoint files", ".pptx")]).replace("/", "\\")
        print(file_path)

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        app.DisplayAlerts = True
        presentation = app.Presentations.Open(FileName=''+file_path)
        presentation.SlideShowSettings.ShowPresenterView = False
    elif presentationType == "gs":
        url = input("Type in the full URL of your Google Slides Presentation: ")
        print("Opening Google Slide URL in standard browser...")
        webbrowser.open(url, new=1, autoraise=True)
        print("\nDisclaimer: For the Motion Controller to work with Google Slides in your webbrowser the browser has to be focussed!")

    cap = cv2.VideoCapture(cam_id)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)

    detector = HandDetector(detectionCon=0.8, maxHands=2)

    action = False
    presentationMode = False

    while True:
        success, img = cap.read()
        hands, img = detector.findHands(img)
        cv2.imshow("Image", img)
        cv2.waitKey(1)

        if len(hands) == 1 or len(hands) == 2:
            if detector.fingersUp(hands[0]) == [1, 1, 1, 1, 1] or (len(hands) == 2 and detector.fingersUp(hands[1]) == [1, 1, 1, 1, 1]):
                if not action:
                    print("start/stop presentation mode")
                    presentationMode = not presentationMode
                    action = True
                    if presentationType == "pp":
                        if presentationMode:
                            presentation.SlideShowSettings.Run()
                        elif not presentationMode:
                            presentation.SlideShowWindow.View.Exit()
                    elif presentationType == "gs":
                        if presentationMode:
                            keyboard.press_and_release("ctrl+f5")
                        elif not presentationMode:
                            keyboard.press_and_release("esc")
            elif detector.fingersUp(hands[0]) == [0, 1, 0, 0, 0] or (len(hands) == 2 and detector.fingersUp(hands[1]) == [0, 1, 0, 0, 0]):
                if not action:
                    print("next slide")
                    action = True
                    if presentationType == "pp":
                        if presentationMode:
                            presentation.SlideShowWindow.View.Next()
                    elif presentationType == "gs":
                        keyboard.press_and_release("right")
            elif detector.fingersUp(hands[0]) == [0, 0, 1, 0, 0] or (len(hands) == 2 and detector.fingersUp(hands[1]) == [0, 0, 1, 0, 0]):
                if not action:
                    print("previous slide")
                    action = True
                    if presentationType == "pp":
                        if presentationMode:
                            presentation.SlideShowWindow.View.Previous()
                    elif presentationType == "gs":
                        keyboard.press_and_release("left")
            elif detector.fingersUp(hands[0]) == [0, 0, 0, 0, 1] or (len(hands) == 2 and detector.fingersUp(hands[1]) == [0, 0, 0, 0, 1]):
                if not action:
                    action = True
                    if presentationType == "pp":
                        app.Quit()
                    elif presentationType == "gs":
                        keyboard.press_and_release("ctrl+w")
                    cv2.destroyAllWindows()
                    exit("Program closed via gesture.")
            else:
                action = False
        else:
            action = False
else:
    print("Invalid input... closing")
    time.sleep(2)
    exit("Invalid input on type selection")

