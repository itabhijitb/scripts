#! /usr/bin/env python

import keyboard
import mouse
import random
import time
from screeninfo import get_monitors
monitor = get_monitors()[0]
width = monitor.width
height = monitor.height
while True:
    ix = random.randint(int(0.1*width), width)
    iy = random.randint(int(0.1*width), height)
    print(ix, iy, mouse.get_position())
    mouse.move(str(ix), str(iy))
    keyboard.press_and_release("ctrl+shift")
    time.sleep(random.randint(5,15))