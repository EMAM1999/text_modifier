import os
import subprocess
import sys
import urllib.request


# helper function
def ensure_package(pkg):
    try:
        __import__(pkg)
    except ImportError:
        print(f"{pkg} not found, installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])


# نتأكد من المكتبات المهمة
# ensure_package("pywin32")
ensure_package("psutil")

# import win32com.client
import psutil

# --------- modifier_core.py content ---------
MODIFIER_CORE_FILE = r"""
import sys
import pyperclip

# -*- coding: utf-8 -*-

# خريطة الحروف العربية <-> الإنجليزية حسب مكانها على الكيبورد
arabic_to_english = {
    "ض": "q",
    "َ": "Q",
    "ص": "w",
    "ً": "W",
    "ث": "e",
    "ُ": "E",
    "ق": "r",
    "ٌ": "R",
    "ف": "t",
    "لإ": "T",
    "غ": "y",
    "إ": "Y",
    "ع": "u",
    "‘": "U",
    "ه": "i",
    "÷": "I",
    "خ": "o",
    "×": "O",
    "ح": "p",
    "؛": "P",
    "ج": "[",
    "<": "{",
    "د": "]",
    ">": "}",
    "ش": "a",
    "ِ": "A",
    "س": "s",
    "ٍ": "S",
    "ي": "d",
    "]": "D",
    "ب": "f",
    "[": "F",
    "ل": "g",
    "لأ": "G",
    "ا": "h",
    "أ": "H",
    "ت": "j",
    "ـ": "J",
    "ن": "k",
    "،": "K",
    "م": "l",
    "/": "L",
    "ك": ";",
    ":": ":",
    "ط": "'",
    "\\": "]",
    "ئ": "z",
    "~": "Z",
    "ء": "x",
    "ْ": "X",
    "ؤ": "c",
    "}": "C",
    "ر": "v",
    "{": "V",
    "لا": "b",
    "لآ": "B",
    "ى": "n",
    "آ": "N",
    "ة": "m",
    "’": "M",
    "و": ",",
    "ز": ".",
    ".": ">",
    "ظ": "/",
    "؟": "?",
    "ذ": "`",
    "ّ": "~",
}

# خريطة الحروف الإنجليزية → العربية
english_to_arabic = {v: k for k, v in arabic_to_english.items()}


def auto_convert(text):
    result = ""
    for ch in text:
        if ch in arabic_to_english:
            result += arabic_to_english[ch]
        elif ch in english_to_arabic:
            result += english_to_arabic[ch]
        else:
            result += ch
    return result


def modify_text(text: str) -> str:
    return auto_convert(text)  # مثال: تحويل النص لحروف كبيرة


text = pyperclip.paste()
pyperclip.copy(modify_text(text))
"""

# --------- AHK Script ---------
AHK_FILE = r"""
^!c:: ; CTRL+ALT+C
{
    Send("^x")
    ClipWait()
    RunWait('python.exe "' A_ScriptDir '\modifier_core.py"', , "Hide")
    ClipWait()
    Send(A_Clipboard)
}
"""

# --------- Uninstaller ---------
UNINSTALLER_FILE = r"""
import os, sys, psutil, shutil

def remove_shortcut():
    startup_folder = os.path.join(os.getenv("APPDATA"), r"Microsoft\Windows\Start Menu\Programs\Startup")
    shortcut_path = os.path.join(startup_folder, "modifier.ahk")
    target_path_py = os.path.join(startup_folder, "modifier_core.py")

    if os.path.exists(shortcut_path) or os.path.exists(target_path_py):
        try:
            os.remove(shortcut_path)
        except:
            pass
        try:
            os.remove(target_path_py)
        except:
            pass
        print("Program removed from Startup.")
    else:
        print("No startup shortcut found.")

def kill_modifier_process():
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if proc.info['name'] and "modifier.exe" in proc.info['name']:
                print(f"Killing process {proc.info['pid']} ({proc.info['name']})")
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

def cleanup_files():
    files_to_delete = ["modifier_core.py", "modifier.ahk", "installer.spec", "modifier.exe", "uninstaller.py", "uninstaller.spec"]
    for f in files_to_delete:
        if os.path.exists("../"+f):
            os.remove(f)
            print(f"Deleted {f}")

if __name__ == "__main__":
    remove_shortcut()
    kill_modifier_process()
    cleanup_files()
    print("Uninstall complete.")
"""


# --------- Installer Functions ---------
def ensure_python():
    try:
        subprocess.run(
            ["python", "--version"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        return "python"
    except:
        print("Python not found. Downloading installer...")
        url = "https://www.python.org/ftp/python/3.12.6/python-3.12.6-amd64.exe"
        installer = "python_installer.exe"
        urllib.request.urlretrieve(url, installer)
        subprocess.run(
            [installer, "/quiet", "InstallAllUsers=1", "PrependPath=1"], check=True
        )
        return "python"


def ensure_pyperclip(python_cmd):
    try:
        subprocess.run(
            [python_cmd, "-m", "pip", "show", "pyperclip"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except:
        print("Installing pyperclip...")
        subprocess.run([python_cmd, "-m", "pip", "install", "pyperclip"], check=True)


def ensure_pyinstaller(python_cmd):
    try:
        subprocess.run(
            [python_cmd, "-m", "pip", "show", "pyinstaller"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except:
        print("Installing pyinstaller...")
        subprocess.run([python_cmd, "-m", "pip", "install", "pyinstaller"], check=True)


def create_files():
    with open("modifier_core.py", "w", encoding="utf-8") as f:
        f.write(MODIFIER_CORE_FILE)
    with open("modifier.ahk", "w", encoding="utf-8") as f:
        f.write(AHK_FILE)
    with open("uninstaller.py", "w", encoding="utf-8") as f:
        f.write(UNINSTALLER_FILE)
    print("Created modifier_core.py, modifier.ahk and uninstaller.py")


# def compile_ahk_to_exe():
#     ahk2exe_path = r"C:\\Program Files\\AutoHotkey\\Compiler\\Ahk2Exe.exe"
#     ahk_script = os.path.abspath("modifier.ahk")
#     output_exe = os.path.abspath("modifier.exe")

#     if os.path.exists(ahk2exe_path):
#         subprocess.run(
#             [ahk2exe_path, "/in", ahk_script, "/out", output_exe], check=True
#         )
#         print(f"Compiled modifier.ahk -> modifier.exe")
#         return output_exe
#     else:
#         print("⚠️ Ahk2Exe not found! Please install AutoHotkey with Compiler.")
#         return None


def build_uninstaller(python_cmd):
    ensure_pyinstaller(python_cmd)
    print("Building uninstaller.exe...")
    subprocess.run(
        [python_cmd, "-m", "PyInstaller", "--onefile", "uninstaller.py"], check=True
    )
    dist_path = os.path.join("dist", "uninstaller.exe")
    if os.path.exists(dist_path):
        target_path = os.path.abspath(dist_path)
        os.replace(dist_path, target_path)
        print(f"Uninstaller.exe created at {target_path}")


def add_to_startup():
    import shutil
    # مسار الـ Startup
    startup_folder = os.path.join(
        os.getenv("APPDATA"), r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    target_path = os.path.join(startup_folder, "modifier.ahk")
    target_path_py = os.path.join(startup_folder, "modifier_core.py")

    # نسخ ملف ahk للـ Startup
    shutil.copy("modifier.ahk", target_path)
    shutil.copy("modifier_core.py", target_path_py)

    print(f"Program added to Startup: {startup_folder}")


# --------- Main ---------
if __name__ == "__main__":
    python_cmd = ensure_python()
    ensure_pyperclip(python_cmd)
    create_files()
    build_uninstaller(python_cmd)

    add_to_startup()
    print(
        "✅ Setup complete! modifier.exe is compiled and will auto-start with Windows."
    )
    # exe_path = compile_ahk_to_exe()
    # if exe_path:
    #     # add_shortcut_to_startup(exe_path)
    #     print(
    #         "✅ Setup complete! modifier.exe is compiled and will auto-start with Windows."
    #     )
    # else:
    #     print(
    #         "❌ Could not compile modifier.ahk to exe. Install AutoHotkey with Ahk2Exe and retry."
    #     )
