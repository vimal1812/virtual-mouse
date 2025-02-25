import cv2
import mediapipe as mp
import math
import win32api
import win32con
import ctypes

mp_drawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands

cap = cv2.VideoCapture(0)
cap.set(3, 1000)
cap.set(4, 600)

hands = mp_hands.Hands(
    max_num_hands=1,
    min_detection_confidence=0.7,
    min_tracking_confidence=0.7)

# Variables for smoothing
smoothing_factor = 0.5  # Adjust the smoothing factor as needed
smoothed_mouse_x = 0
smoothed_mouse_y = 0

# Variables for drag and drop
dragging = False
drag_start_x = 0
drag_start_y = 0

# Scroll speed (adjust as needed)
scroll_speed = 3

# Constants for keycodes
VK_LWIN = 0x5B
VK_TAB = 0x09

# Load user32.dll
user32 = ctypes.windll.user32

while True:
    success, image = cap.read()

    if not success:
        break

    image = cv2.flip(image, 1)

    # Convert the image from BGR to RGB
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

    # To improve performance, optionally mark the image as not writeable to
    # pass by reference.
    image.flags.writeable = False
    results = hands.process(image)

    # Draw the hand annotations on the image.
    image.flags.writeable = True
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(
                image,
                hand_landmarks,
                mp_hands.HAND_CONNECTIONS)

            # Get the landmarks of the thumb, index finger, middle finger, ring finger, and pinky
            thumb = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
            index_finger = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
            middle_finger = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
            ring_finger = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
            pinky = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]

            # Move the mouse cursor to the position of the index finger
            screen_width, screen_height = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
            target_mouse_x = int(index_finger.x * screen_width)
            target_mouse_y = int(index_finger.y * screen_height)

            # Apply exponential smoothing to the cursor position
            smoothed_mouse_x = (1 - smoothing_factor) * smoothed_mouse_x + smoothing_factor * target_mouse_x
            smoothed_mouse_y = (1 - smoothing_factor) * smoothed_mouse_y + smoothing_factor * target_mouse_y

            # Round the smoothed cursor position to integers
            mouse_x = int(smoothed_mouse_x)
            mouse_y = int(smoothed_mouse_y)

            win32api.SetCursorPos((mouse_x, mouse_y))

            # Check if the thumb and index finger are close together for left-click
            distance_thumb_index = math.dist([thumb.x, thumb.y], [index_finger.x, index_finger.y])
            if distance_thumb_index < 0.05:
                # Perform a left-click operation with the mouse
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, mouse_x, mouse_y, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, mouse_x, mouse_y, 0, 0)


            # Check if the index finger and middle finger are close together for right-click
            distance_index_middle = math.dist([index_finger.x, index_finger.y], [middle_finger.x, middle_finger.y])
            if distance_index_middle < 0.03:
                # Perform a right-click operation with the mouse
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, mouse_x, mouse_y, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, mouse_x, mouse_y, 0, 0)

            # Check if the thumb and middle finger are close together for drag and drop
            distance_thumb_middle = math.dist([thumb.x, thumb.y], [middle_finger.x, middle_finger.y])
            if distance_thumb_middle < 0.05:
                if not dragging:
                    # Start dragging by performing a left-click and hold
                    dragging = True
                    click_start_x = mouse_x
                    click_start_y = mouse_y
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, mouse_x, mouse_y, 0, 0)
                else:
                    # Continue dragging by moving the mouse
                    win32api.SetCursorPos((mouse_x, mouse_y))

            else:
                if dragging:
                    # End dragging by releasing the left mouse button
                    dragging = False
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, mouse_x, mouse_y, 0, 0)



            # Check if the thumb and pinky are close together for scroll up
            distance_thumb_pinky = math.dist([thumb.x, thumb.y], [pinky.x, pinky.y])
            if distance_thumb_pinky < 0.05:
                # Perform a scroll up operation
                win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, mouse_x, mouse_y, 120, 0)


            # Check if the thumb and ring finger are close together for scroll down
            distance_thumb_ring = math.dist([thumb.x, thumb.y], [ring_finger.x, ring_finger.y])
            if distance_thumb_ring < 0.05:
                # Perform a scroll down operation
                win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, mouse_x, mouse_y, -120, 0)


            # Check if the thumb, index finger, and middle finger are close together for task view
            distance_thumb_index_middle = math.dist([thumb.x, thumb.y], [index_finger.x, index_finger.y])
            distance_index_middle = math.dist([index_finger.x, index_finger.y], [middle_finger.x, middle_finger.y])
            if distance_thumb_index_middle < 0.05 and distance_index_middle < 0.05:
                # Perform the task view action by simulating Win + Tab key combination
                user32.keybd_event(VK_LWIN, 0, 0, 0)  # Press Win key
                user32.keybd_event(VK_TAB, 0, 0, 0)  # Press Tab key
                user32.keybd_event(VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release Tab key
                user32.keybd_event(VK_LWIN, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release Win key

    cv2.imshow('Hand Tracking', image)
    if cv2.waitKey(5) & 0xFF == 27:
        break

hands.close()
cap.release()
cv2.destroyAllWindows()
