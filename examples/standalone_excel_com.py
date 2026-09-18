"""Windows Excel COM ownership and paths. No spreadsheet library dependency."""
from pathlib import Path
from contextlib import contextmanager
import gc, json, os, time
import pythoncom, win32com.client, win32process

ROOT = Path(__file__).resolve().parent.parent
WORK = ROOT / 'work'
OUTPUT = ROOT / 'output'
SOURCES = ['August through YE.xlsx', 'Cinci Submission 9102027.xlsx', 'Revenue Tracker - 2026 09 16.xlsx']

@contextmanager
def excel_session():
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    settings = {'DisplayAlerts': app.DisplayAlerts, 'EnableEvents': app.EnableEvents,
                'ScreenUpdating': app.ScreenUpdating, 'AskToUpdateLinks': app.AskToUpdateLinks}
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.ScreenUpdating = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    print(json.dumps({'excel_version':app.Version, 'owned_pid':pid}), flush=True)
    try:
        yield app
    finally:
        try:
            for book in list(app.Workbooks):
                book.Close(SaveChanges=False)
            for key,value in settings.items():
                setattr(app,key,value)
            app.Quit()
        finally:
            app = None
            gc.collect()
            pythoncom.CoUninitialize()
        print(json.dumps({'closed_owned_excel_pid':pid}), flush=True)

def colname(n):
    s=''
    while n:
        n,r=divmod(n-1,26);s=chr(65+r)+s
    return s

def matrix(value):
    if isinstance(value,tuple): return [list(r) if isinstance(r,tuple) else [r] for r in value]
    return [[value]]

