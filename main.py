import math
import time
from multiprocessing import Process, Queue
import tkinter as tk
import queue

import cv2
import mediapipe as mp
import pyautogui
import pygetwindow as gw
import win32con
import win32gui
import screen_brightness_control as sbc

# --------------------------------------------------------------------------------
# --- Глобальні змінні та налаштування ---
# --------------------------------------------------------------------------------
pyautogui.FAILSAFE = False
DEAD_ZONE_RADIUS = 35
JOYSTICK_SENSITIVITY = 0.3
PINCH_THRESHOLD = 0.05
CLICK_COOLDOWN = 1.0


class GestureState:
    def __init__(self):
        self.last_click_time = 0
        self.active_special_gesture = None

        self.is_active = False
        self.last_mode_change_time = 0
        self.MODE_COOLDOWN = 1.0

        self.swipe_motion_ready = False
        self.swipe_motion_start_x = 0.0
        self.swipe_motion_start_y = 0.0
        self.last_swipe_motion_time = 0.0

        self.SWIPE_MOTION_THRESHOLD_X = 0.1
        self.SWIPE_MOTION_THRESHOLD_Y = 0.1

        self.SWIPE_MOTION_COOLDOWN = 0.8

        self.swipe_motion_in_grace_period = False
        self.swipe_motion_lost_time = 0.0
        self.SWIPE_GRACE_PERIOD = 0.25

        self.volume_mode = False
        self.scroll_mode = False
        self.brightness_mode = False

        self.mode_anchor_y = 0.0

        self.app_context = 'general'
        self.last_context_check_time = 0
        self.CONTEXT_CHECK_COOLDOWN = 1.0

        self.last_action_time = 0
        self.ACTION_COOLDOWN = 0.05


# --------------------------------------------------------------------------------
# --- ДОПОМІЖНІ ФУНКЦІЇ ---
# --------------------------------------------------------------------------------

def count_fingers_up(hand_landmarks, hand_label):
    fingers_up = []

    thumb_tip_x = hand_landmarks[4][0]
    thumb_mcp_x = hand_landmarks[2][0]

    if hand_label == 'Right':
        if thumb_tip_x > thumb_mcp_x:
            fingers_up.append(1)
        else:
            fingers_up.append(0)
    elif hand_label == 'Left':
        if thumb_tip_x < thumb_mcp_x:
            fingers_up.append(1)
        else:
            fingers_up.append(0)

    tip_ids = [8, 12, 16, 20]
    pip_ids = [6, 10, 14, 18]
    for tip_id, pip_id in zip(tip_ids, pip_ids):
        if hand_landmarks[tip_id][1] < hand_landmarks[pip_id][1]:
            fingers_up.append(1)
        else:
            fingers_up.append(0)

    return fingers_up


def is_scissors_gesture(fingers_up_list):
    return fingers_up_list[1] and fingers_up_list[2] and not fingers_up_list[3] and not fingers_up_list[4]


def is_flat_palm_gesture(fingers_up_list):
    return fingers_up_list.count(1) >= 4


def is_fist(fingers_up_list):
    return fingers_up_list.count(1) == 0


def is_thumbs_up(fingers_up_list):
    if fingers_up_list is None or len(fingers_up_list) < 3:
        return False

    is_thumb_ok = fingers_up_list[0] == 1
    is_index_down = fingers_up_list[1] == 0
    is_middle_down = fingers_up_list[2] == 0

    return is_thumb_ok and is_index_down and is_middle_down


def is_thumbs_down(hand_landmarks, fingers_up_list):
    if fingers_up_list is None or len(fingers_up_list) < 5:
        return False

    other_fingers_down = not (fingers_up_list[1] or fingers_up_list[2] or fingers_up_list[3] or fingers_up_list[4])
    if not other_fingers_down:
        return False

    thumb_tip_y = hand_landmarks[4][1]
    thumb_pip_y = hand_landmarks[3][1]

    is_thumb_pointing_down = (thumb_tip_y > thumb_pip_y)

    return is_thumb_pointing_down


def is_three_fingers(fingers_up_list):
    if fingers_up_list is None: return False
    return fingers_up_list == [0, 1, 1, 1, 0]


def is_v_sign(fingers_up_list):
    if fingers_up_list is None: return False
    return fingers_up_list[1] and fingers_up_list[2] and not fingers_up_list[3] and not fingers_up_list[4]


def is_pinky_up(fingers_up_list):
    if fingers_up_list is None: return False
    return fingers_up_list == [0, 0, 0, 0, 1]


def is_ok_gesture(hand_coords, fingers_up_list):
    if fingers_up_list is None: return False
    thumb_tip = hand_coords[4]
    index_tip = hand_coords[8]
    dist_thumb_index = math.hypot(thumb_tip[0] - index_tip[0], thumb_tip[1] - index_tip[1])
    is_touching = dist_thumb_index < (PINCH_THRESHOLD * 1.5)
    other_fingers_up = fingers_up_list[2] and fingers_up_list[3] and fingers_up_list[4]
    return is_touching and other_fingers_up


# --------------------------------------------------------------------------------
# --- РОБОЧИЙ ПРОЦЕС 1: ДЕТЕКЦІЯ РУК ---
# --------------------------------------------------------------------------------
def detection_worker(frame_queue, gesture_queue):
    hands = mp.solutions.hands.Hands(
        model_complexity=0, min_detection_confidence=0.6, min_tracking_confidence=0.5, max_num_hands=2)
    DETECTION_SCALE_FACTOR = 0.5
    print("Detection worker started...")
    while True:
        frame = frame_queue.get()
        if frame is None: break
        h, w, _ = frame.shape
        small_frame = cv2.resize(frame, (int(w * DETECTION_SCALE_FACTOR), int(h * DETECTION_SCALE_FACTOR)),
                                 interpolation=cv2.INTER_AREA)
        rgb_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
        result = hands.process(rgb_frame)
        simplified_hands = []
        if result.multi_hand_landmarks and result.multi_handedness:
            for hand_landmarks, handedness in zip(result.multi_hand_landmarks, result.multi_handedness):
                hand_coords = [(lm.x, lm.y, lm.z) for lm in hand_landmarks.landmark]
                hand_label = handedness.classification[0].label
                simplified_hands.append((hand_coords, hand_label, hand_landmarks))
        gesture_queue.put((frame, simplified_hands, result.multi_hand_landmarks))
    hands.close()
    print("Detection worker stopped.")


# --------------------------------------------------------------------------------
# --- РОБОЧИЙ ПРОЦЕС 3: ЛОГІКА ЖЕСТІВ ---
# --------------------------------------------------------------------------------
def gesture_worker(gesture_queue, display_queue, command_queue, gui_queue):
    state = GestureState()

    print("Gesture worker started...")
    prev_time = 0

    def send_gui_update(text):
        try:
            while not gui_queue.empty():
                gui_queue.get_nowait()
            gui_queue.put_nowait(text)
        except queue.Empty:
            pass
        except Exception:
            pass

    def send_command(command):
        try:
            command_queue.put_nowait(command)
        except:
            pass

    while True:
        frame, simplified_hands, raw_landmarks = gesture_queue.get()
        if frame is None: break

        h, w, _ = frame.shape
        state.active_special_gesture = None
        current_time = time.time()
        can_change_mode = (current_time - state.last_mode_change_time) > state.MODE_COOLDOWN

        profile_action_taken = False
        action_text = ""
        status_text = ""
        fingers_up = None
        hand_landmarks = None
        avg_knuckle_x = 0.0
        avg_knuckle_y = 0.0

        if len(simplified_hands) == 2:
            hand1_coords, hand1_label, _ = simplified_hands[0]
            hand2_coords, hand2_label, _ = simplified_hands[1]

            fingers1 = count_fingers_up(hand1_coords, hand1_label)
            fingers2 = count_fingers_up(hand2_coords, hand2_label)

            both_palms = is_flat_palm_gesture(fingers1) and is_flat_palm_gesture(fingers2)
            both_ok = is_ok_gesture(hand1_coords, fingers1) and is_ok_gesture(hand2_coords, fingers2)

            if both_palms and can_change_mode:
                state.is_active = True
                state.last_mode_change_time = current_time
                time.sleep(0.5)
            elif both_ok and state.is_active and can_change_mode:
                state.is_active = False
                state.last_mode_change_time = current_time
                time.sleep(0.5)
            else:
                both_scissors = is_scissors_gesture(fingers1) and is_scissors_gesture(fingers2)
                if both_scissors:
                    state.active_special_gesture = 'scissors'

            profile_action_taken = True


        elif len(simplified_hands) == 1 and state.is_active:

            hand_landmarks, hand_label, _ = simplified_hands[0]
            fingers_up = count_fingers_up(hand_landmarks, hand_label)

            knuckle_landmarks = [hand_landmarks[5], hand_landmarks[9], hand_landmarks[13], hand_landmarks[17]]
            avg_knuckle_x = sum(lm[0] for lm in knuckle_landmarks) / 4.0
            avg_knuckle_y = sum(lm[1] for lm in knuckle_landmarks) / 4.0

            if (current_time - state.last_context_check_time) > state.CONTEXT_CHECK_COOLDOWN:
                state.last_context_check_time = current_time
                active_title = ""
                try:
                    active_window = gw.getActiveWindow()
                    if active_window is not None:
                        active_title = active_window.title.lower()
                        if 'powerpoint' in active_title:
                            state.app_context = 'powerpoint'
                        elif 'zoom' in active_title:
                            state.app_context = 'zoom'
                        elif any(browser in active_title for browser in
                                 ['chrome', 'firefox', 'edge', 'brave', 'нова вкладка', 'opera']):
                            state.app_context = 'browser'
                        elif any(media in active_title for media in ['spotify', 'vlc']):
                            state.app_context = 'media'
                        else:
                            state.app_context = 'general'
                    else:
                        state.app_context = 'general'
                except Exception as e:
                    state.app_context = 'general'

                print(f"DEBUG: Title='{active_title}' | Profile='{state.app_context}'")

            ppt_override = False
            if state.app_context == 'powerpoint':
                if is_flat_palm_gesture(fingers_up):
                    send_command("ppt:start_show")
                    action_text = "Start Slideshow"
                    profile_action_taken = True
                    ppt_override = True

            if not ppt_override:
                is_swipe_gest = is_flat_palm_gesture(fingers_up)

                if is_swipe_gest:
                    profile_action_taken = True
                    state.swipe_motion_in_grace_period = False

                    if not state.swipe_motion_ready:
                        state.swipe_motion_ready = True
                        state.swipe_motion_start_x = avg_knuckle_x
                        state.swipe_motion_start_y = avg_knuckle_y
                        state.last_swipe_motion_time = current_time
                        action_text = "Swipe Armed (Palm)"
                        print("DEBUG: Swipe Armed. Anchor set.")

                    else:
                        delta_x = avg_knuckle_x - state.swipe_motion_start_x
                        delta_y = avg_knuckle_y - state.swipe_motion_start_y

                        can_swipe_now = (current_time - state.last_swipe_motion_time) > state.SWIPE_MOTION_COOLDOWN

                        abs_dx = abs(delta_x)
                        abs_dy = abs(delta_y)
                        is_strong_enough = (abs_dx > state.SWIPE_MOTION_THRESHOLD_X) or (
                                abs_dy > state.SWIPE_MOTION_THRESHOLD_Y)

                        if is_strong_enough and can_swipe_now:

                            if abs_dx > abs_dy:
                                if delta_x > 0:
                                    send_command("swipe:next_window")
                                    action_text = "Swipe Right"
                                    print("DEBUG: Swipe Right complete.")
                                else:
                                    send_command("swipe:prev_window")
                                    action_text = "Swipe Left"
                                    print("DEBUG: Swipe Left complete.")
                            else:

                                if delta_y > 0:
                                    send_command("swipe:desktop")
                                    action_text = "Show Desktop"
                                    print("DEBUG: Swipe Down (Desktop) complete.")
                                else:
                                    send_command("swipe:task_view")
                                    action_text = "Task View"
                                    print("DEBUG: Swipe Up (Task View) complete.")

                            state.swipe_motion_start_x = avg_knuckle_x
                            state.swipe_motion_start_y = avg_knuckle_y
                            state.last_swipe_motion_time = current_time
                            state.last_click_time = current_time

                        elif not action_text:
                            action_text = "Swipe Ready..."

                else:
                    if state.swipe_motion_ready:
                        if not state.swipe_motion_in_grace_period:
                            state.swipe_motion_in_grace_period = True
                            state.swipe_motion_lost_time = current_time
                            profile_action_taken = True
                            action_text = "Swipe Ready..."
                        else:
                            if (current_time - state.swipe_motion_lost_time) > state.SWIPE_GRACE_PERIOD:
                                print("DEBUG: Swipe Disarmed (Grace period ended).")
                                state.swipe_motion_ready = False
                                state.swipe_motion_in_grace_period = False
                            else:
                                profile_action_taken = True
                                action_text = "Swipe Ready..."

                    else:
                        state.swipe_motion_in_grace_period = False

            if not profile_action_taken:
                status_text = f"MODE: {state.app_context.upper()}"
                can_click = (current_time - state.last_click_time) > CLICK_COOLDOWN

                if fingers_up == [0, 1, 0, 0, 0]:
                    center_x, center_y = w / 2, h / 2
                    index_x_norm, index_y_norm, _ = hand_landmarks[8]
                    index_x_abs, index_y_abs = int(index_x_norm * w), int(index_y_norm * h)
                    move_x = index_x_abs - center_x
                    move_y = index_y_abs - center_y
                    if math.hypot(move_x, move_y) > DEAD_ZONE_RADIUS:
                        send_command(f"move:{move_x},{move_y}")
                    action_text = "Cursor Mode"
                    cv2.line(frame, (int(center_x), int(center_y)), (index_x_abs, index_y_abs), (0, 255, 0), 2)
                    profile_action_taken = True

                elif state.volume_mode:
                    if not is_v_sign(fingers_up):
                        state.volume_mode = False
                    else:
                        current_y = hand_landmarks[9][1]
                        delta_y = current_y - state.mode_anchor_y
                        can_act = (current_time - state.last_action_time) > state.ACTION_COOLDOWN
                        if delta_y < -0.04 and can_act:
                            send_command("vol_up")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Volume Up"
                        elif delta_y > 0.04 and can_act:
                            send_command("vol_down")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Volume Down"
                        else:
                            action_text = "VOLUME MODE"
                        profile_action_taken = True

                elif state.scroll_mode:
                    if not is_thumbs_up(fingers_up):
                        state.scroll_mode = False
                    else:
                        current_y = hand_landmarks[9][1]
                        delta_y = current_y - state.mode_anchor_y
                        can_act = (current_time - state.last_action_time) > state.ACTION_COOLDOWN
                        if delta_y < -0.04 and can_act:
                            send_command("scroll_up")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Scroll Up"
                        elif delta_y > 0.04 and can_act:
                            send_command("scroll_down")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Scroll Down"
                        else:
                            action_text = "SCROLL MODE"
                        profile_action_taken = True

                elif state.brightness_mode:
                    if not is_pinky_up(fingers_up):
                        state.brightness_mode = False
                    else:
                        current_y = hand_landmarks[20][1]
                        delta_y = current_y - state.mode_anchor_y
                        can_act = (current_time - state.last_action_time) > state.ACTION_COOLDOWN

                        if delta_y < -0.04 and can_act:
                            send_command("brightness_up")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Brightness ++"
                        elif delta_y > 0.04 and can_act:
                            send_command("brightness_down")
                            state.mode_anchor_y = current_y
                            state.last_action_time = current_time
                            action_text = "Brightness --"
                        else:
                            action_text = "BRIGHTNESS MODE"
                        profile_action_taken = True

                if not profile_action_taken:

                    if can_click:
                        thumb_tip = hand_landmarks[4]
                        index_tip = hand_landmarks[8]
                        dist_thumb_index = math.hypot(thumb_tip[0] - index_tip[0], thumb_tip[1] - index_tip[1])
                        is_left_click = dist_thumb_index < PINCH_THRESHOLD
                        is_right_click = (fingers_up == [1, 1, 0, 0, 1])

                        if state.app_context == 'general':
                            if is_left_click:
                                send_command("click")
                                action_text = "Left Click"
                                profile_action_taken = True
                            elif is_right_click:
                                send_command("right_click")
                                action_text = "Right Click"
                                profile_action_taken = True

                        elif state.app_context == 'powerpoint':
                            if is_three_fingers(fingers_up):
                                send_command("ppt:next_slide")
                                action_text = "Next Slide"
                                profile_action_taken = True
                            elif is_fist(fingers_up):
                                send_command("ppt:prev_slide")
                                action_text = "Previous Slide"
                                profile_action_taken = True
                            elif is_left_click:
                                send_command("click")
                                action_text = "Left Click"
                                profile_action_taken = True
                            elif is_right_click:
                                send_command("right_click")
                                action_text = "Right Click"
                                profile_action_taken = True

                        elif state.app_context == 'zoom':
                            if is_v_sign(fingers_up):
                                send_command("zoom:mute")
                                action_text = "Mute/Unmute"
                                profile_action_taken = True
                            elif is_three_fingers(fingers_up):
                                send_command("zoom:video")
                                action_text = "Start/Stop Video"
                                profile_action_taken = True
                            elif is_left_click:
                                send_command("click")
                                action_text = "Left Click"
                                profile_action_taken = True
                            elif is_right_click:
                                send_command("right_click")
                                action_text = "Right Click"
                                profile_action_taken = True

                        elif state.app_context == 'browser':
                            if is_thumbs_down(hand_landmarks, fingers_up):
                                send_command("browser:prev_tab")
                                action_text = "Previous Tab"
                                profile_action_taken = True
                            elif is_three_fingers(fingers_up):
                                send_command("browser:next_tab")
                                action_text = "Next Tab"
                                profile_action_taken = True
                            elif is_left_click:
                                send_command("click")
                                action_text = "Left Click"
                                profile_action_taken = True
                            elif is_right_click:
                                send_command("right_click")
                                action_text = "Right Click"
                                profile_action_taken = True

                        elif state.app_context == 'media':
                            if is_three_fingers(fingers_up):
                                send_command("media:next_track")
                                action_text = "Next Track"
                                profile_action_taken = True
                            elif is_fist(fingers_up):
                                send_command("media:prev_track")
                                action_text = "Prev Track"
                                profile_action_taken = True
                            elif is_left_click:
                                send_command("click")
                                action_text = "Left Click"
                                profile_action_taken = True
                            elif is_right_click:
                                send_command("right_click")
                                action_text = "Right Click"
                                profile_action_taken = True

                        if not profile_action_taken:
                            if is_v_sign(fingers_up):
                                state.volume_mode = True
                                state.mode_anchor_y = hand_landmarks[9][1]
                                state.last_action_time = current_time
                                action_text = "Volume Mode ENGAGED"
                                profile_action_taken = True

                            elif is_thumbs_up(fingers_up):
                                state.scroll_mode = True
                                state.mode_anchor_y = hand_landmarks[9][1]
                                state.last_action_time = current_time
                                action_text = "Scroll Mode ENGAGED"
                                profile_action_taken = True

                            elif is_pinky_up(fingers_up):
                                state.brightness_mode = True
                                state.mode_anchor_y = hand_landmarks[20][1]
                                state.last_action_time = current_time
                                action_text = "Brightness Mode ENGAGED"
                                profile_action_taken = True

                        if profile_action_taken:
                            state.last_click_time = current_time

        else:
            if not state.is_active:
                status_text = "MODE: INACTIVE"
            else:
                status_text = f"MODE: {state.app_context.upper()}"

            if state.swipe_motion_ready:
                print("DEBUG: Swipe Disarmed (Hand lost).")
            state.swipe_motion_ready = False

        if state.active_special_gesture is not None:
            pass

        cv2.circle(frame, (int(w / 2), int(h / 2)), DEAD_ZONE_RADIUS, (0, 255, 0), 2)

        if (current_time - prev_time) > 0:
            fps = 1 / (current_time - prev_time)
            prev_time = current_time
            cv2.putText(frame, f'FPS: {int(fps)}', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

        if not state.is_active:
            cv2.putText(frame, "(Show TWO PALMS to activate)", (10, 110), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
        else:
            cv2.putText(frame, "(Show TWO 'OK' to deactivate)", (10, 110), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0),
                        2)

        if action_text:
            send_gui_update(action_text)
        else:
            send_gui_update(status_text)

        display_queue.put((frame, raw_landmarks))

    print("Gesture worker stopped.")
    gui_queue.put(None)


# --------------------------------------------------------------------------------
# --- РОБОЧИЙ ПРОЦЕС 4: ВИКОНАННЯ ДІЙ ---
# --------------------------------------------------------------------------------
def action_worker(command_queue):
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0

    print("Action worker started...")

    while True:
        command = command_queue.get()
        if command is None: break

        if command.startswith("move:"):
            try:
                _, coords = command.split(':')
                move_x, move_y = map(float, coords.split(','))
                pyautogui.move(move_x * JOYSTICK_SENSITIVITY, move_y * JOYSTICK_SENSITIVITY)
            except Exception as e:
                print(f"Action worker: Помилка парсингу руху: {e}")

        elif command == "click":
            pyautogui.click()
        elif command == "right_click":
            pyautogui.rightClick()

        elif command == "scroll_up":
            pyautogui.scroll(120)
        elif command == "scroll_down":
            pyautogui.scroll(-120)

        elif command == "vol_up":
            pyautogui.press('volumeup')
        elif command == "vol_down":
            pyautogui.press('volumedown')

        elif command == "brightness_up":
            try:
                current = sbc.get_brightness()
                level = current[0] if isinstance(current, list) else current
                sbc.set_brightness(min(100, level + 5))
            except Exception as e:
                print(f"Brightness error: {e}")
        elif command == "brightness_down":
            try:
                current = sbc.get_brightness()
                level = current[0] if isinstance(current, list) else current
                sbc.set_brightness(max(0, level - 5))
            except Exception as e:
                print(f"Brightness error: {e}")

        elif command == "swipe:next_window":
            pyautogui.hotkey('alt', 'tab')
        elif command == "swipe:prev_window":
            pyautogui.hotkey('alt', 'shift', 'tab')
        elif command == "swipe:desktop":
            pyautogui.hotkey('win', 'd')
        elif command == "swipe:task_view":
            pyautogui.hotkey('win', 'tab')

        elif command == "ppt:start_show":
            pyautogui.press('f5')
        elif command == "ppt:next_slide":
            pyautogui.press('right')
        elif command == "ppt:prev_slide":
            pyautogui.press('left')

        elif command == "zoom:raise_hand":
            pyautogui.hotkey('alt', 'y')
        elif command == "zoom:mute":
            pyautogui.hotkey('alt', 'a')
        elif command == "zoom:video":
            pyautogui.hotkey('alt', 'v')

        elif command == "browser:next_tab":
            pyautogui.hotkey('ctrl', 'tab')
        elif command == "browser:prev_tab":
            pyautogui.hotkey('ctrl', 'shift', 'tab')

        elif command == "media:play_pause":
            pyautogui.press('space')
        elif command == "media:next_track":
            pyautogui.press('nexttrack')
        elif command == "media:prev_track":
            pyautogui.press('prevtrack')

    print("Action worker stopped.")


# --------------------------------------------------------------------------------
# --- РОБОЧИЙ ПРОЦЕС 5: GUI ОВЕРЛЕЙ ---
# --------------------------------------------------------------------------------
def gui_worker(gui_queue):
    try:
        root = tk.Tk()
        root.title("Gesture Status")
        root.geometry("300x50+50+50")

        root.wm_attributes("-topmost", True)
        root.wm_attributes("-alpha", 0.7)
        root.config(bg='black')
        root.overrideredirect(True)

        status_label = tk.Label(
            root,
            text="Initializing...",
            font=("Arial", 16, "bold"),
            fg="cyan",
            bg="black"
        )
        status_label.pack(expand=True, fill="both")

        def check_queue():
            try:
                while not gui_queue.empty():
                    message = gui_queue.get_nowait()
                    if message is None:
                        root.destroy()
                        return
                    status_label.config(text=message)
            except queue.Empty:
                pass
            root.after(50, check_queue)

        print("GUI worker started...")
        root.after(50, check_queue)
        root.mainloop()

    except Exception as e:
        print(f"GUI Error: {e}")
    finally:
        print("GUI worker stopped.")


# --------------------------------------------------------------------------------
# --- ГОЛОВНИЙ ПРОЦЕС (КАМЕРА + ВІДОБРАЖЕННЯ) ---
# --------------------------------------------------------------------------------
if __name__ == '__main__':
    print("Main process started...")
    print("IMPORTANT: Make sure you have installed 'pywin32' (pip install pywin32)")
    print("IMPORTANT: Make sure you have installed 'pygetwindow' (pip install pygetwindow)")

    mp_draw = mp.solutions.drawing_utils
    mp_hands = mp.solutions.hands

    frame_queue = Queue(maxsize=1)
    gesture_queue = Queue(maxsize=1)
    display_queue = Queue(maxsize=1)
    command_queue = Queue(maxsize=5)
    gui_queue = Queue(maxsize=5)

    detection_process = Process(target=detection_worker, args=(frame_queue, gesture_queue))
    gesture_process = Process(target=gesture_worker, args=(gesture_queue, display_queue, command_queue, gui_queue))
    action_process = Process(target=action_worker, args=(command_queue,))
    gui_process = Process(target=gui_worker, args=(gui_queue,))

    detection_process.daemon = True
    gesture_process.daemon = True
    action_process.daemon = True
    gui_process.daemon = True

    detection_process.start()
    gesture_process.start()
    action_process.start()
    gui_process.start()

    cap = cv2.VideoCapture(1)
    if not cap.isOpened():
        print("ПОМИЛКА: Не вдалося відкрити камеру. Перевірте індекс (можливо, 0?)")
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("ПОМИЛКА: Камеру з індексом 0 також не знайдено.")
            frame_queue.put(None)
            gesture_queue.put(None)
            command_queue.put(None)
            gui_queue.put(None)
            exit()

    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
    cap.set(cv2.CAP_PROP_FPS, 30)

    final_frame_to_show = None
    raw_landmarks_to_draw = None

    window_name = "Gesture Control (5-Core Pipeline)"
    window_set_top = False

    while True:
        success, frame = cap.read()
        if not success:
            print("Помилка читання кадру, спроба перепідключення...")
            cap.release()
            cap = cv2.VideoCapture(1)
            time.sleep(1)
            continue

        frame = cv2.flip(frame, 1)

        try:
            frame_queue.put_nowait(frame)
        except:
            pass
        try:
            final_frame_to_show, raw_landmarks_to_draw = display_queue.get_nowait()
        except:
            pass

        if final_frame_to_show is not None:
            if raw_landmarks_to_draw:
                for hand_landmarks in raw_landmarks_to_draw:
                    mp_draw.draw_landmarks(final_frame_to_show, hand_landmarks, mp_hands.HAND_CONNECTIONS,
                                           mp_draw.DrawingSpec(color=(255, 255, 0), thickness=2, circle_radius=2),
                                           mp_draw.DrawingSpec(color=(255, 0, 255), thickness=2))

            cv2.imshow(window_name, final_frame_to_show)

        if not window_set_top and final_frame_to_show is not None:
            try:
                hwnd = win32gui.FindWindow(None, window_name)
                if hwnd:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                    print("--- Вікно камери встановлено 'Always on Top' ---")
                    window_set_top = True
            except Exception as e:
                print(f"Помилка встановлення 'Always on Top': {e}")
                window_set_top = True

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    print("Shutting down...")
    frame_queue.put(None)
    gesture_queue.put(None)
    command_queue.put(None)
    gui_queue.put(None)

    detection_process.join()
    gesture_process.join()
    action_process.join()
    gui_process.join()

    cap.release()
    cv2.destroyAllWindows()
    print("Main process finished.")