import cv2 as cv
import numpy as np
import os
import pyautogui
import time
import win32gui, win32ui, win32con
from PIL import ImageGrab



os.chdir(os.path.dirname(os.path.abspath(__file__)))

# properties
delay_between_commands = 1.00

windowtocapture = 'EVE - '
window_handle = win32gui.FindWindow(None, windowtocapture)
if not window_handle:
    raise Exception('Window not found: {}'.format(windowtocapture))
region = win32gui.GetWindowRect(window_handle)

# dimensions of the window that GetWindowRect collects, used for other calculations 
top_right_corner = region[2], region[1]
top_left_corner = region[0], region[1]
bottom_left_corner = region[0], region[3]
bottom_right_corner = region[2], region[3]

# using the regions gathered, work out the location of the overview for more precise image identification
overview_bottom_right = ((bottom_right_corner[0] - 10), (bottom_right_corner[1] - 325))
overview_bottom_left = ((overview_bottom_right[0] - 390), (overview_bottom_right[1] - 0))
overview_top_right = ((top_right_corner[0] - 10), (top_right_corner[1] + 150))
overview_top_left = ((overview_top_right[0] - 390), (overview_top_right[1] + 0))
overview_region = (overview_top_left[0], overview_top_left[1], overview_bottom_right[0], overview_bottom_right[1])


def initialisedPyautoGUI():    
        pyautogui.FAILSAFE = True
        print("FAILSAFE = True")

def countdownTimer():
    #Countdown timer
    print("Starting", end="")
    for i in range(0, 5):
        print(".", end="")
        time.sleep(1)
    print("Go")

def holdkey(key, seconds=1.00):
    pyautogui.keyDown(key)
    time.sleep(seconds)
    pyautogui.keyUp(key)
    time.sleep(delay_between_commands)

def warptokeepstarinMJ():
    warp_to_keepstar_MJ = pyautogui.locateOnScreen(r'images\warp_to_keepstar_MJ.png', region, confidence= 0.7)
    pyautogui.moveTo(warp_to_keepstar_MJ)
    pyautogui.click(warp_to_keepstar_MJ)
    warpto()
    pyautogui.click(warp_to_keepstar_MJ), print("Warping to the Keepstar in MJ.")

def pulsemwd():
    pyautogui.keyDown('alt')
    pyautogui.keyDown('1')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('1')

def warptosun():
    warptosun = pyautogui.locateOnScreen(r'images\Sun_Icon.png', region, confidence= 0.9)
    pyautogui.moveTo(warptosun)
    pyautogui.click(warptosun)
    warpto()
    pyautogui.click(warptosun), print("Warping to the Sun.")

def warpto():
    holdkey('w', 1.00)
    pyautogui.click(), print("I'm entering warp.")

def approach():
    holdkey('q')
    pyautogui.click()

def dockedcheck():
    dockcheck = pyautogui.locateOnScreen(r'images\undock.png', region, confidence= 0.9)
    if dockcheck is None:
        print("I'm not docked.")
    else:
        print("I'm undocking my ship...")
        pyautogui.moveTo(dockcheck, None, 1)
        pyautogui.click(dockcheck)
        time.sleep(8)
        stopship()
        print("Stopping the ship.")
        time.sleep(3)
        print("Centering camera")
        orbitcamera = pyautogui.locateOnScreen(r'images\orbit_camera_icon.png', region, confidence= 0.8)
        pyautogui.moveTo(orbitcamera, None, 1)
        pyautogui.click(clicks=2, interval= 0.25)

def stopship():
    pyautogui.keyDown('ctrl'), pyautogui.keyDown('space')
    pyautogui.keyUp('ctrl'), pyautogui.keyUp('space')

def inwarpwait():
   while True:
    if  pyautogui.locateOnScreen(r'images\warp_drive_active.png', region, confidence= 0.9) is None:
        print("I'm no longer in warp.")
        time.sleep(1)
        break
    else:
        pyautogui.locateOnScreen(r'images\warp_drive_active.png', region, confidence= 0.9)
        print("I'm in warp...")
        time.sleep(1)
    
def list_window_names():
        def winEnumHandler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                print(hex(hwnd),'"' + win32gui.GetWindowText(hwnd) + '"')
        win32gui.EnumWindows(winEnumHandler, None)

def Startup():

    initialisedPyautoGUI()
    # start countdown timer
    countdownTimer()
    
def warptofortinRQ():
    warp_to_fort_RQ = pyautogui.locateOnScreen(r'images\warp_to_fort_RQ.png', region, confidence= 0.7)
    pyautogui.moveTo(warp_to_fort_RQ)
    pyautogui.click(warp_to_fort_RQ)
    warpto()
    pyautogui.click(warp_to_fort_RQ), print("Warping to the Fortizar in RQOO.")

def openlocations():
    Notification_icon = pyautogui.locateOnScreen(r'images\notification_button.png', region, confidence= 0.7)
    pyautogui.moveTo(Notification_icon,  None, 1)
    pyautogui.move(30, 0, .25)
    pyautogui.rightClick()
    Location_rightclickmenu = pyautogui.locateOnScreen(r'images\right_click_menu_locations_icon.png', region, confidence= 0.9)
    pyautogui.moveTo(Location_rightclickmenu, None, 0.5)

def warptoPI1():
    pi1_button = pyautogui.locateOnScreen(r'images\PI#1_button.png', region, confidence= 0.8)
    pyautogui.moveTo(pi1_button,  None, 1)
    pyautogui.click()
    pulsemwd(), print("pulsing MWD for faster warp")
    time.sleep(.5)
    pulsemwd()

def accesscustomsoffice():
        pi_tab = pyautogui.locateOnScreen(r'images\PI_Tab_overview.png', region, confidence= 0.9)
        pyautogui.moveTo(pi_tab,  None, 1)
        pyautogui.leftClick(), print("Opening PI tab.")
        pyautogui.move(0,  50, .5)
        pyautogui.rightClick()
        pyautogui.move(0,  210, .25)
        pyautogui.leftClick(), print("Accessing Customs Office.")
        
def mainactions():

    # this is to undock the ship, warp it to bookmarks made above each planet customs office, access the 
    # customs office, unload it into the fleet hangar in the DST, warp to the next customs office,
    # do this 6 times in total, then warp to the fortizar and dock.

    Startup()
    dockedcheck()
    print("Beginning PI run...")
    time.sleep(2)
    openlocations()
    print("Opening locations via right click drop down menu.")
    time.sleep(.5)
    warptoPI1()
    print("Warping to bookmark PI #1")
    time.sleep(30.0)
    inwarpwait()
    time.sleep(1.0)
    accesscustomsoffice()


    #time.sleep(5.00)
    #inwarpwait()
    #time.sleep(1.00)
    #warptoPI1()
    # Set of tasks to do in this order first
    #dockedcheck()
    #time.sleep(1.00)
    #warptosun()
    #time.sleep(5.00)
    #inwarpwait()
    #time.sleep(1.00)
    #warptokeepstarinMJ()
    #time.sleep(5.00)
    #inwarpwait()

    
   
       
mainactions()



    
