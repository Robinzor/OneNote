import win32com.client
import pytz
import datetime


ids = [i for i in range(100)]

for i in ids:
    try:
        OneNote_AppID = win32com.client.Dispatch(f'OneNote.Application.{i}')
        print(OneNote_AppID)
    except:
        pass

