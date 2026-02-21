import time
import os
import ctypes
import configparser

import keyboard
import pyperclip
import pystray
from PIL import Image


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_INPUTLANGCHANGEREQUEST = 0x0050
KLF_ACTIVATE = 0x00000001
SMTO_ABORTIFHUNG = 0x0002

HKL_EN = "00000409"  # English (US)
HKL_RU = "00000419"  # Russian


# Same physical key maps to different characters depending on layout
EN_TO_RU = {
    "q": "й", "w": "ц", "e": "у", "r": "к", "t": "е", "y": "н",
    "u": "г", "i": "ш", "o": "щ", "p": "з", "[": "х", "]": "ъ",
    "a": "ф", "s": "ы", "d": "в", "f": "а", "g": "п", "h": "р",
    "j": "о", "k": "л", "l": "д", ";": "ж", "'": "э",
    "z": "я", "x": "ч", "c": "с", "v": "м", "b": "и", "n": "т",
    "m": "ь", ",": "б", ".": "ю", "`": "ё",

    "Q": "Й", "W": "Ц", "E": "У", "R": "К", "T": "Е", "Y": "Н",
    "U": "Г", "I": "Ш", "O": "Щ", "P": "З", "{": "Х", "}": "Ъ",
    "A": "Ф", "S": "Ы", "D": "В", "F": "А", "G": "П", "H": "Р",
    "J": "О", "K": "Л", "L": "Д", ":": "Ж", '"': "Э",
    "Z": "Я", "X": "Ч", "C": "С", "V": "М", "B": "И", "N": "Т",
    "M": "Ь", "<": "Б", ">": "Ю", "~": "Ё",
}
RU_TO_EN = {v: k for k, v in EN_TO_RU.items()}


def sleep_short(seconds=0.06):
    time.sleep(seconds)


class ClipboardGuard:
    """Save/restore clipboard safely."""
    def __enter__(self):
        try:
            self._saved = pyperclip.paste()
        except Exception:
            self._saved = ""
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            pyperclip.copy(self._saved)
        except Exception:
            pass


def detect_language(text: str) -> str:
    """Return 'ru' if mostly Cyrillic, else 'en'."""
    ru_count = sum(1 for c in text if "\u0400" <= c <= "\u04FF")
    en_count = sum(1 for c in text if c.isascii() and c.isalpha())
    return "ru" if ru_count >= en_count else "en"


def convert_text(text: str) -> str:
    lang = detect_language(text)
    if lang == "ru":
        return "".join(RU_TO_EN.get(c, c) for c in text)
    return "".join(EN_TO_RU.get(c, c) for c in text)


def _copy_selection() -> str:
    """
    Try Ctrl+C and return clipboard text.
    Assumes clipboard is already cleared by caller for detection.
    """
    keyboard.send("ctrl+c")
    sleep_short(0.12)
    try:
        return pyperclip.paste()
    except Exception:
        return ""


def _select_previous_word_and_copy() -> str:
    """
    More reliable "last word":
    - clears selection
    - selects previous word (Ctrl+Shift+Left)
    - copies
    """
    pyperclip.copy("")
    sleep_short(0.03)

    # First try selecting the previous word
    keyboard.send("ctrl+shift+left")
    sleep_short(0.06)
    txt = _copy_selection()

    print (f"Selected previous word: '{txt}'")  # Debugging output

    # Some apps behave differently; fallback to selecting next word if needed
    if not txt.strip():
        keyboard.send("right")  # collapse selection/caret
        sleep_short(0.03)
        keyboard.send("ctrl+shift+right")
        sleep_short(0.06)
        txt = _copy_selection()

    return txt.strip()


def _get_foreground_hwnd() -> int:
    return user32.GetForegroundWindow()


def _send_lang_change(hwnd: int, hkl: int) -> None:
    """
    Try to request language change in a way that works for more apps.
    """
    # Try focused control first (can be 0 for some apps)
    focused = user32.GetFocus()
    target = focused if focused else hwnd

    # SendMessageTimeout reduces "hung window" issues
    user32.SendMessageTimeoutW(
        target,
        WM_INPUTLANGCHANGEREQUEST,
        0,
        hkl,
        SMTO_ABORTIFHUNG,
        200,
        ctypes.byref(ctypes.c_ulong())
    )


def get_current_lang() -> str:
    hwnd = _get_foreground_hwnd()
    thread_id = user32.GetWindowThreadProcessId(hwnd, None)
    hkl = user32.GetKeyboardLayout(thread_id)
    langid = hkl & 0xFFFF
    return "ru" if langid == 0x0419 else "en"


def switch_keyboard_layout(to_lang: str) -> None:
    layout_id = HKL_EN if to_lang == "en" else HKL_RU
    hwnd = _get_foreground_hwnd()
    if not hwnd:
        return

    hkl = user32.LoadKeyboardLayoutW(layout_id, KLF_ACTIVATE)
    if not hkl:
        return

    _send_lang_change(hwnd, hkl)


def on_hotkey():
    with ClipboardGuard():
        pyperclip.copy("")
        sleep_short(0.04)

        selected = _copy_selection()

        if not selected.strip():
            current_lang = get_current_lang()
            switch_keyboard_layout("en" if current_lang == "ru" else "ru")
            return

        trimmed = selected.strip()
        original_lang = detect_language(trimmed)
        converted = convert_text(trimmed)

        pyperclip.copy(converted)
        sleep_short(0.04)
        keyboard.send("ctrl+v")
        sleep_short(0.06)

        switch_keyboard_layout("en" if original_lang == "ru" else "ru")


def load_config():
    config = configparser.ConfigParser()
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
    config.read(config_path, encoding="utf-8")
    return config.get("settings", "hotkey", fallback="F16")


def create_tray_image() -> Image.Image:
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
    return Image.open(icon_path)


def run_tray(hotkey: str) -> None:
    def on_exit(icon, item):
        keyboard.unhook_all()
        icon.stop()

    menu = pystray.Menu(
        pystray.MenuItem(f"Hotkey: {hotkey}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit", on_exit),
    )
    icon = pystray.Icon("k5enru", create_tray_image(), "EN↔RU Converter", menu)
    icon.run()


def main():
    hotkey = load_config()
    print("EN<->RU layout converter running.")
    print(f"Hotkey: {hotkey}")
    print("Right-click tray icon to exit.\n")

    keyboard.add_hotkey(hotkey, on_hotkey, suppress=True)
    run_tray(hotkey)


if __name__ == "__main__":
    main()
