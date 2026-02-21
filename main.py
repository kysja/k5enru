import sys
import time
import os
import ctypes
import configparser

import pyperclip
import pystray
from PIL import Image
from pynput import keyboard as pynput_kb


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_INPUTLANGCHANGEREQUEST = 0x0050
KLF_ACTIVATE = 0x00000001
SMTO_ABORTIFHUNG = 0x0002

HKL_EN = "00000409"  # English (US)
HKL_RU = "00000419"  # Russian

kb = pynput_kb.Controller()
hotkey_listener: pynput_kb.GlobalHotKeys | None = None


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


def copy_selection() -> str:
    with kb.pressed(pynput_kb.Key.ctrl):
        kb.press('c')
        kb.release('c')
    sleep_short(0.12)
    try:
        return pyperclip.paste()
    except Exception:
        return ""


def _get_foreground_hwnd() -> int:
    return user32.GetForegroundWindow()


def _send_lang_change(hwnd: int, hkl: int) -> None:
    focused = user32.GetFocus()
    target = focused if focused else hwnd
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

        selected = copy_selection()

        if not selected.strip():
            current_lang = get_current_lang()
            switch_keyboard_layout("en" if current_lang == "ru" else "ru")
            return

        trimmed = selected.strip()
        original_lang = detect_language(trimmed)
        converted = convert_text(trimmed)

        pyperclip.copy(converted)
        sleep_short(0.04)
        with kb.pressed(pynput_kb.Key.ctrl):
            kb.press('v')
            kb.release('v')
        sleep_short(0.06)

        switch_keyboard_layout("en" if original_lang == "ru" else "ru")


def to_pynput_hotkey(hotkey_str: str) -> str:
    """Convert 'F16' or 'ctrl+shift+f' to pynput GlobalHotKeys format."""
    key_names = {k.name for k in pynput_kb.Key}
    parts = hotkey_str.lower().strip().split('+')
    result = []
    for part in parts:
        part = part.strip()
        if part in key_names:
            result.append(f'<{part}>')
        else:
            result.append(part)
    return '+'.join(result)


def get_resource_path(filename: str) -> str:
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def create_tray_image() -> Image.Image:
    return Image.open(get_resource_path("icon.png"))


def run_tray(hotkey: str) -> None:
    def on_exit(icon, item):
        if hotkey_listener:
            hotkey_listener.stop()
        icon.stop()

    menu = pystray.Menu(
        pystray.MenuItem(f"Hotkey: {hotkey}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit", on_exit),
    )
    icon = pystray.Icon("k5enru", create_tray_image(), "EN↔RU Converter", menu)
    icon.run()


def main():
    global hotkey_listener

    hotkey = load_config()
    pynput_key = to_pynput_hotkey(hotkey)
    print("EN<->RU layout converter running.")
    print(f"Hotkey: {hotkey} ({pynput_key})")
    print("Right-click tray icon to exit.\n")

    hotkey_listener = pynput_kb.GlobalHotKeys({pynput_key: on_hotkey})
    hotkey_listener.start()
    run_tray(hotkey)


def load_config():
    config = configparser.ConfigParser()
    config.read(get_resource_path("config.ini"), encoding="utf-8")
    return config.get("settings", "hotkey", fallback="F16")


if __name__ == "__main__":
    main()
