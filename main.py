# Import necessary libraries
import cv2  # For image processing and webcam operations
import mediapipe as mp  # For hand tracking
import math  # For mathematical operations
import numpy as np  # For numerical operations
import pyautogui  # For simulating keyboard presses
from ctypes import cast, POINTER  # For casting pointers in the pycaw library
from comtypes import CLSCTX_ALL  # For COM library context
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # For volume control

# Initialize MediaPipe solutions for drawing and hand tracking
mp_drawing = mp.solutions.drawing_utils
mp_drawing_styles = mp.solutions.drawing_styles
mp_hands = mp.solutions.hands

# Initialize volume control using Pycaw library
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volRange = volume.GetVolumeRange()
minVol, maxVol, volBar, volPer = volRange[0], volRange[1], 400, 0

# Setup the webcam
wCam, hCam = 800, 640
cam = cv2.VideoCapture(0)
cam.set(3, wCam)  # Set width
cam.set(4, hCam)  # Set height

# Flag to indicate if play/pause has been activated
play_pause_executed = False

# Setup MediaPipe hand model
with mp_hands.Hands(
    model_complexity=0,
    min_detection_confidence=0.5,
    min_tracking_confidence=0.5) as hands:

    # Main loop to continuously capture images from the webcam
    while cam.isOpened():
        success, image = cam.read()  # Read an image from the webcam
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)  # Convert the image from BGR to RGB
        results = hands.process(image)  # Process the image for hand tracking
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)  # Convert back to BGR for OpenCV operations

        # Check for two hands and toggle play/pause
        if results.multi_hand_landmarks and len(results.multi_hand_landmarks) == 2:
            if not play_pause_executed:
                pyautogui.press('playpause')  # Simulate Play/Pause key press
                play_pause_executed = True
        else:
            play_pause_executed = False  # Reset the flag when the condition is not met

        # Draw the hand landmarks
        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_drawing.draw_landmarks(
                    image,
                    hand_landmarks,
                    mp_hands.HAND_CONNECTIONS,
                    mp_drawing_styles.get_default_hand_landmarks_style(),
                    mp_drawing_styles.get_default_hand_connections_style()
                )

        # Initialize a list to store landmark positions
        lmList = []
        if results.multi_hand_landmarks:
            myHand = results.multi_hand_landmarks[0]
            for id, lm in enumerate(myHand.landmark):
                h, w, c = image.shape  # Get the height, width, and channels of the image
                cx, cy = int(lm.x * w), int(lm.y * h)  # Calculate the x and y coordinates
                lmList.append([id, cx, cy])  # Append the coordinates to the list

        # Volume control based on the distance between thumb and index finger
        if len(lmList) != 0:
            x1, y1 = lmList[4][1], lmList[4][2]  # Position of the thumb
            x2, y2 = lmList[8][1], lmList[8][2]  # Position of the index finger
            cv2.circle(image, (x1, y1), 15, (255, 255, 255))  # Draw a circle on the thumb
            cv2.circle(image, (x2, y2), 15, (255, 255, 255))  # Draw a circle on the index finger
            cv2.line(image, (x1, y1), (x2, y2), (0, 255, 0), 3)  # Draw a line between thumb and index finger
            length = math.hypot(x2 - x1, y2 - y1)  # Calculate the distance between thumb and index finger

            # Interpolate the length to volume range
            vol = np.interp(length, [50, 220], [minVol, maxVol])
            volume.SetMasterVolumeLevel(vol, None)  # Set the volume
            volBar = np.interp(length, [50, 220], [400, 150])  # Interpolate for the volume bar
            volPer = np.interp(length, [50, 220], [0, 100])  # Interpolate for the volume percentage

            # Draw the volume bar and percentage
            cv2.rectangle(image, (50, 150), (85, 400), (0, 0, 0), 3)
            cv2.rectangle(image, (50, int(volBar)), (85, 400), (0, 0, 0), cv2.FILLED)
            if play_pause_executed:
                cv2.putText(image, 'Play/Pause', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 0), 3)
            else:
                cv2.putText(image, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 0), 3)

        # Display the image with the hand tracking and volume information
        cv2.imshow('Media Control Hand Gesture', image)
        if cv2.waitKey(1) & 0xFF == ord('q'):  # Exit the loop if 'q' is pressed
            break

# Release the webcam and destroy all OpenCV windows
cam.release()
cv2.destroyAllWindows()
