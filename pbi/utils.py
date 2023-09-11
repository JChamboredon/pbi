import ctypes
import subprocess
import sys


def run_ps_script(ps_script):
    p = subprocess.Popen(
        [
            "powershell.exe",
            ps_script
        ],
        stdout=sys.stdout
    )
    p.communicate()


def message_box(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)