# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2017 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import random
import re
import logging
import threading
import win32con

from lib.common.abstracts import Auxiliary
from lib.common.defines import (
    KERNEL32, USER32, WM_GETTEXT, WM_GETTEXTLENGTH, WM_CLOSE, BM_CLICK,
    EnumWindowsProc, EnumChildProc, create_unicode_buffer
)

log = logging.getLogger(__name__)


def foreach_child(hwnd, lparam):
    # List of buttons labels to click.
    buttons = [
        "yes", "oui",
        "ok",
        "i accept",
        "next", "suivant",
        "new", "nouveau",
        "install", "installer",
        "file", "fichier",
        "run", "start", "marrer", "cuter",
        "i agree", "accepte",
        "enable", "activer", "accord", "valider",
        "don't send", "ne pas envoyer",
        "don't save",
        "continue", "continuer",
        "personal", "personnel",
        "scan", "scanner",
        "unzip", "dezip",
        "open", "ouvrir",
        "close the program",
        "execute", "executer",
        "launch", "lancer",
        "save", "sauvegarder",
        "download", "load", "charger",
        "end", "fin", "terminer"
        "later",
        "finish",
        "end",
        "allow access",
        "remind me later",
        "save", "sauvegarder",
        "update"
    ]

    # List of buttons labels to not click.
    dontclick = [
        "don't run",
        "i do not accept",
        "check for a solution and close the program",
        "close the program",
        "never allow opening files of this type",
        "always allow opening files of this type"
    ]

    classname = create_unicode_buffer(50)
    USER32.GetClassNameW(hwnd, classname, 50)

    # Check if the class of the child is button.
    if "button" in classname.value.lower():
        # Get the text of the button.
        length = USER32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
        text = create_unicode_buffer(length + 1)
        USER32.SendMessageW(hwnd, WM_GETTEXT, length + 1, text)

        # Check if the button is set as "clickable" and click it.
        textval = text.value.replace("&", "").lower()
        for button in buttons:
            if button in textval:
                for btn in dontclick:
                    if btn in textval:
                        break
                else:
                    log.info("Found button %r, clicking it" % text.value)
                    USER32.SetForegroundWindow(hwnd)
                    KERNEL32.Sleep(1000)
                    USER32.SendMessageW(hwnd, BM_CLICK, 0, 0)

    # Recursively search for childs (USER32.EnumChildWindows).
    return True


def get_window_text(hwnd):
    text = create_unicode_buffer(1024)
    USER32.GetWindowTextW(hwnd, text, 1024)
    return text.value


def get_office_window(hwnd, lparam):
    '''
    Callback procedure invoked for every enumerated window.
    Purpose is to close any office window.
    '''
    if USER32.IsWindowVisible(hwnd):
        text = get_window_text(hwnd)
        if re.search("(Microsoft|Word|Excel|PowerPoint)", text):
            USER32.SendNotifyMessageW(hwnd, WM_CLOSE, None, None)
            KERNEL32.Sleep(1000)
            log.info("Closed Office window: %s", text)
    return True


def foreach_window(hwnd, lparam):
    '''Callback procedure invoked for every enumerated window. 
    '''

    # If the window is visible, enumerate its child objects, looking
    # for buttons.
    if USER32.IsWindowVisible(hwnd):
        USER32.EnumChildWindows(hwnd, EnumChildProc(foreach_child), 0)
    return True


def move_mouse(x, y):
    USER32.SetCursorPos(x, y)


def click_mouse(x, y):
    USER32.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    USER32.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def double_click(x, y):
    click_mouse(x, y)
    KERNEL32.Sleep(50)
    click_mouse(x, y)


def set_full_screen(hwnd):
    SW_MAXIMISE = 3
    USER32.ShowWindow(hwnd, SW_MAXIMISE)
    KERNEL32.Sleep(100)


class Coordinates(object):
    '''window coordinates helper class. '''

    X_JUMPS = 55
    Y_JUMPS = 60

    def __init__(self, x_padding, y_padding):
        """
        Parameters:
                y_padding, number representing the boundaries in Y axis
        """
        self.y_padding = y_padding
        self.x_padding = x_padding
        self.max_x = int(USER32.GetSystemMetrics(0) - x_padding)
        self.max_y = int(USER32.GetSystemMetrics(1) - 100)
        self._generator = self._cords_generator()

    def _cords_generator(self):
        '''This generator will yeild next (x, y) cords of divided screen matrix. '''
        current_x, current_y = self.x_padding, self.y_padding

        while current_y < self.max_y:
            yield current_x, current_y

            if current_x >= self.max_x - self.x_padding:
                current_x = self.x_padding
                current_y += self.Y_JUMPS
            else:
                current_x += self.X_JUMPS

    def next(self):
        '''Retrun next cords of screen matrix. '''
        try:
            x, y = next(self._generator)
        except:
            x, y = self.center()
        return x, y

    def center(self):
        '''Return center screen cords.'''
        return self.max_x / 2, self.max_y / 2

    def random(self):
        '''Return random (x, y) cords. '''
        return random.randint(0, self.max_x), random.randint(0, self.max_y)


class Human(threading.Thread, Auxiliary):
    """Human after all"""

    def __init__(self, options={}, analyzer=None):
        threading.Thread.__init__(self)
        Auxiliary.__init__(self, options, analyzer)
        self.do_run = True
        self.parse_options()
        self.coordinates = Coordinates(170, 300)

    def parse_options(self):
        # Global disable flag.
        if "human" in self.options:
            self.do_move_mouse = int(self.options["human"])
            self.do_click_mouse = int(self.options["human"])
            self.do_click_buttons = int(self.options["human"])
        else:
            self.do_move_mouse = True
            self.do_click_mouse = True
            self.do_click_buttons = True

        # Per-feature enable or disable flag.
        if "human.move_mouse" in self.options:
            self.do_move_mouse = int(self.options["human.move_mouse"])

        if "human.click_mouse" in self.options:
            self.do_click_mouse = int(self.options["human.click_mouse"])

        if "human.click_buttons" in self.options:
            self.do_click_buttons = int(self.options["human.click_buttons"])

    def stop(self):
        self.do_run = False

    def run(self):
        # human starts before the sample invocation, wait for 8s to start
        minimal_timeout = KERNEL32.GetTickCount() + 2000
        # set office close timeout after 2/3 of analysis (in milliseconds)
        office_close_sec = int(self.options.get("timeout") * (3. / 4) * 1000)
        office_close_timeout = KERNEL32.GetTickCount() + office_close_sec
        is_office_close = False
        is_full_screen = False
        is_ultrafast = self.options.get("timeout") == 25
        # adaptive sleep timer
        sleep = 50 if is_ultrafast else 1000

        while self.do_run:

            KERNEL32.Sleep(sleep)  # we wait for minimal timeout anyway so no loss here

            if KERNEL32.GetTickCount() < minimal_timeout:
                continue

            if not is_office_close and KERNEL32.GetTickCount() > office_close_timeout:
                USER32.EnumWindows(EnumWindowsProc(get_office_window), 0)
                is_office_close = True

            if self.do_click_mouse and self.do_move_mouse:
                # extract foregroud window name
                fg_window_name = ""
                hwnd = USER32.GetForegroundWindow()
                try:
                    fg_window_name = get_window_text(hwnd).lower()
                except:
                    log.exception("failed to extract window name")
                    pass

                # make the office window on front
                if fg_window_name == "":
                    x, y = self.coordinates.center()
                    click_mouse(x, y)

                if "word" in fg_window_name or "excel" in fg_window_name:
                    if not is_full_screen:
                        set_full_screen(hwnd)
                        is_full_screen = True
                    x, y = self.coordinates.next()
                    move_mouse(x, y)
                    double_click(x, y)

                if not is_ultrafast:
                    # make random move
                    x, y = self.coordinates.random()
                    move_mouse(x, y)
                    # click_mouse(x, y)

            if self.do_click_buttons:
                USER32.EnumWindows(EnumWindowsProc(foreach_window), 0)
