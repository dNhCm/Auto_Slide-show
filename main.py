
# Imports
import json
import os
from datetime import datetime
from time import sleep
from threading import Thread

import pyautogui as ag


# Global variables
ROOT = "\\".join(os.path.abspath(__file__).split('\\')[:-1])
DELAY = int(open(ROOT + '\\data\\delay.txt', 'r').read())
schedule = json.load(open(ROOT + '\\data\\schedule.json', 'r'))


# Threads

def start_slide_show():
    path = ROOT + f"\\media"
    os.startfile(path)
    ag.sleep(DELAY)
    ag.hotkey("ctrl", "a")
    ag.sleep(DELAY)
    ag.press("enter")
    ag.sleep(DELAY)
    ag.press("f5")


def shutdown():
    """
    Thread that control shutdown of Windows system

    :return:
    """
    # Get shutdown datetime
    now = datetime.now()
    shutdown_time = schedule[now.weekday()].split(':')
    shutdown_date = datetime(
        year=now.year,
        month=now.month,
        day=now.day,
        hour=int(shutdown_time[0]),
        minute=int(shutdown_time[1])
    )

    # Wait for shutdown
    delta = shutdown_date - now
    sleep(delta.seconds)

    # Shutdown
    path = "C:\\Windows\\System32\\shutdown.bat"  # from OP
    os.startfile(path)


# Main

def main():
    slide_show_thread = Thread(target=start_slide_show)
    shutdown_thread = Thread(target=shutdown)
    slide_show_thread.start()
    shutdown_thread.start()


if __name__ == "__main__":
    main()
