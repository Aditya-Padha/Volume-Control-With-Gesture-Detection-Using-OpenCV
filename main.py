#install opencv
import math
import cv2
import numpy
import mediapipe as mp
from ctypes import cast, POINTER

import numpy as np
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
vrange = volume.GetVolumeRange()
minv, maxv = vrange[0], vrange[1]
#volume.SetMasterVolumeLevel(-20.0, None)

mp_hands = mp.solutions.hands
draw = mp.solutions.drawing_utils
hands = mp_hands.Hands()

capture = cv2.VideoCapture(0)
while True:
    value, image = capture.read()
    rgbimage = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    processed_image = hands.process(rgbimage)
    print(processed_image.multi_hand_landmarks)
    if processed_image.multi_hand_landmarks:
        for handlandmarks in processed_image.multi_hand_landmarks:
            for finger_id, landmark_co in enumerate(handlandmarks.landmark):
                #print(finger_id, landmark_co)
                height, width, channel = image.shape
                cx, cy = int(landmark_co.x * width), int(landmark_co.y * height)
                #print(finger_id,cx,cy)
                if finger_id == 4:
                    cv2.circle(image, (cx, cy), 20, (255, 0, 255), cv2.FILLED)
                    tpx, tpy = cx, cy
                if finger_id == 8:
                    cv2.circle(image, (cx, cy), 20, (255, 0, 255), cv2.FILLED)
                    ipx, ipy = cx, cy
                    cv2.line(image, (tpx, tpy), (ipx, ipy), (0, 255, 0), 9)
                    distance = math.sqrt((ipx-tpx)**2 + (ipy-tpy)**2)
                    print(distance)
                    v = np.interp(distance, [25, 250], [minv, maxv])
                    volume.SetMasterVolumeLevel(v, None)

            draw.draw_landmarks(image, handlandmarks, mp_hands.HAND_CONNECTIONS)

    cv2.imshow('Image Capture', image)
    if cv2.waitKey(1) & 0xFF == 27:
        break
capture.release()
cv2.destroyAllWindows()

