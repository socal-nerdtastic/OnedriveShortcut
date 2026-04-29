#!/usr/bin/env python3
#
#  Copyright 2026 j.brauer <j.brauer@bruker.com>
#
#  This program is internal property of Bruker Nano, and may not be distributed.
#

import html
from urllib.parse import urlsplit, unquote
import json
from pathlib import Path
from dataclasses import dataclass
import subprocess
import sys

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from makeshortcut import winmakeshortcut

CONFIG = Path.home() / ".OnedriveShortcuts.json"

@dataclass
class Settings:
    # ~ sp_link:str = ""
    exe_path:str = ""
    save_loc:str = ""
    ex_map:dict = None

DEFAULT_MAP = {
    ('xls', 'xlt', 'xlm', 'xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xla', 'xlam', 'xll', 'xlw'): 'EXCEL.EXE',
    ('doc', 'dot', 'wbk', 'docx', 'dotx', 'docm', 'dotm'): 'WINWORD.EXE',
    ('ppt', 'pot', 'pps', 'ppa', 'pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm', 'sldx', 'sldm'): 'POWERPNT.EXE',
    ('accda', 'accdb', 'accde', 'accdr', 'accdt', 'accdu'): 'MSACCESS.EXE',
    ('one',): 'ONENOTE.EXE',
    ('msg',): 'OUTLOOK.EXE',
}
EXE_SEARCH_LOCATIONS:list[Path] = [
    Path(r"C:\Program Files (x86)"), # 32-bit programs
    Path(r"C:\Program Files"), # 64-bit programs
]

def search_for_exe(exename):
    # optimization would be nice.
    for fold in EXE_SEARCH_LOCATIONS:
        for root_fold in fold.rglob('root'):
            if root_fold.is_dir():
                for fn in root_fold.rglob(exename):
                    return (fn)

def build_map():
    result = {}
    for exts, progname in DEFAULT_MAP.items():
        if (exe_path := search_for_exe(progname)):
            result |= dict.fromkeys(exts, exe_path)
        else:
            print(progname, "NOT FOUND!")
    return result

try:
    with open(CONFIG) as f:
        data = json.load(f)
except Exception as e:
    print("Failed to load settings file:", type(e).__name__, e)
    data = dict(ex_map = build_map())
sett = Settings(**data) # evil global; ah well

def save_sett():
    with open(CONFIG, 'w') as f:
        json.dump(sett, f, indent=2)

# stolen from: https://github.com/socal-nerdtastic/PyShortcut
def powershell_run(cmd):
    proc = subprocess.run(["powershell", cmd], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    return proc.stdout.decode().strip(), proc.stderr.decode().strip()

class FileCreationError(Exception): pass

def make_lnk(target:str|Path, location:str|Path, description:str=None, working_dir:str=None, arguments:str=None, icon_location:str=None, window_style:str|int=None, hot_key:str=None):
    # https://learn.microsoft.com/en-us/troubleshoot/windows-client/admin-development/create-desktop-shortcut-with-wsh
    if not str(location).endswith((".lnk", ".url")):
        raise FileCreationError("The shortcut pathname must end with .lnk or .url.")
    pscmd = f"$s=(New-Object -COM WScript.Shell).CreateShortcut('{location}');"
    pscmd += rf"$s.TargetPath='{target}';"
    if description:
        pscmd += rf"$s.Description='{description}';"
    if working_dir:
        pscmd += rf"$s.WorkingDirectory='{working_dir}';"
    if arguments:
        pscmd += rf"""$s.Arguments='"{arguments}"';"""
    if icon_location:
        pscmd += rf"$s.IconLocation='{icon_location}';"
    if window_style: # 1=normal; 3=maximized; 7=minimized
        pscmd += rf"$s.WindowStyle='{window_style}';"
    if hot_key:
        pscmd += rf"$s.Hotkey ='{hot_key}';"
    pscmd += r"$s.Save()"
    stdout, stderr = powershell_run(pscmd)
    if stderr:
        print(stderr)
        raise FileCreationError("Could not create shortcut file.")

def get_specialfolder(specialname):
    # https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolder?view=net-10.0
    cmd = f"[Environment]::GetFolderPath([Environment+SpecialFolder]::{specialname})"
    stdout, stderr = powershell_run(cmd)
    if stderr:
        print("ERROR when looking for special folder", specialname)
    return stdout
# end thieving

def cleanup(lnktext):
    lnk_url = urlsplit(lnktext)
    rem_q = lnk_url._replace(query='')
    new_lnk = rem_q.geturl()
    return unquote(new_lnk)

class BrowseFrame(ttk.LabelFrame):
    def __init__(self, parent, cmd=filedialog.askopenfilename, **kwargs):
        self.var = kwargs.pop("textvariable")
        if self.var is None:
            self.var = tk.StringVar()
        super().__init__(parent, **kwargs)
        self.cmd = cmd
        ent = ttk.Entry(self, textvariable=self.var, width=70)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn = ttk.Button(self, text="...", width=3, command=self.on_browse)
        btn.pack(side=tk.RIGHT)
    def on_browse(self):
        fn = self.cmd()
        if not fn:
            return
        self.var.set(fn)

class Main(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        frm = ttk.LabelFrame(self, text="Paste in sharepoint link:")
        frm.pack(fill=tk.X, expand=True)
        self.sp_link_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=self.sp_link_var)
        ent.pack(fill=tk.X, expand=True)
        ent.focus()
        self.sp_link_var.trace_add("write", self.on_link_change)

        self.exe_path_var = tk.StringVar(value=sett.exe_path)
        frm = BrowseFrame(self,text="Executable to use:", textvariable=self.exe_path_var)
        frm.pack(fill=tk.X, expand=True)

        # ~ self.save_loc_var = tk.StringVar(value=sett.save_loc)
        self.save_loc_var = tk.StringVar(value=r"C:\Users\j.brauer\Downloads\EXPORTS")
        frm = BrowseFrame(self,
            text="Save location for the new shortcut:",
            textvariable=self.save_loc_var,
            cmd=filedialog.askdirectory)
        frm.pack(fill=tk.X, expand=True)

        self.status = tk.StringVar()
        status = tk.Label(self, textvariable=self.status)
        status.pack(fill=tk.X)

        frm = ttk.Frame(self)
        frm.pack()
        btn = ttk.Button(frm, text="Cancel", command=self.quit)
        btn.pack(side=tk.LEFT)
        btn = ttk.Button(frm, text="Make Shortcut!", command=self.on_ok)
        btn.pack(side=tk.LEFT)

    def on_close(self):
        sett["sp_link"] = ''
        sett["exe_path"] = ''
        sett["save_loc"] = ''
        save_sett()

    def on_ok(self):
        if not (exeloc := self.exe_path_var.get()):
            return print('no exe')
        if not (cleanlink := cleanup(self.sp_link_var.get())):
            return print('no link')
        if not (save_loc := self.save_loc_var.get()):
            return print('no save loc')
        if not save_loc.endswith(".lnk"):
            return print('link location must end with .lnk')
        make_lnk(target=exeloc, location=save_loc, arguments=cleanlink)

    def make_targetpath(self):
        targetpath = f'"{self.exe_path_var.get()}" "{cleanup(self.sp_link_var.get())}"'
        print((targetpath))
        return targetpath

    def on_link_change(self, event=None, *args):
        print('new link')
        cleanlink = cleanup(self.sp_link_var.get())
        print(repr(cleanlink))
        filetype = cleanlink.rsplit(".",1)[1]
        print(filetype)
        if (exeloc := sett.ex_map.get(filetype)):
            self.exe_path_var.set(exeloc)
            self.status.set(f"Target path len is {len(self.make_targetpath())} characters (limit: 260)")

def main():
    root = tk.Tk()
    win = Main(root)
    win.pack(fill=tk.X, expand=True)
    # ~ root.geometry("600x200")
    root.mainloop()
    # ~ print(root.geometry())

main()

# ~ from pprint import pprint
# ~ x = build_map()
# ~ pprint(x)
