import win32com.client
import pytz
import datetime

C

ids = [i for i in range(100)]
installations = []
for i in ids:
    try:
        OneNote_Find_AppID = win32com.client.Dispatch(f'OneNote.Application.{i}')
        OneNote = OneNote_Find_AppID
        installations.append(OneNote)
    except:
        pass

print(installations)

onObj = win32com.client.gencache.EnsureDispatch('OneNote.Application.12')
result = onObj.GetHierarchy("",win32com.client.constants.hsNotebooks)
print(result)