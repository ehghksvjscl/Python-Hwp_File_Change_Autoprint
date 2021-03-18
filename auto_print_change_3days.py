import os
import win32com.client as win32
import win32event
import win32process
from datetime import datetime, timedelta
from win32com.shell.shell import ShellExecuteEx
from win32com.shell import shellcon 
from pytimekr import pytimekr # 한국 공유일 리스트

# global vars
path = 'C:\\auto_print\\3days.hwp'
now = datetime.now()

def setting_hwp(path):
    # hwp settings
    hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule('FilePathCheckDLL','SecurityModule')
    hwp.Open(path,"HWP","forceopen:true")

    return hwp

def get_date(hwp,kr_holiday,path,now):
    days = get_day_counts(now,kr_holiday)
    page_index = 0
    for i in days:
        get_field_in_hwp(hwp,i,page_index)
        page_index += 1

    hwp.Clear(3)
    hwp.Quit()
    ahto_print(path)

# Get day counts
def get_day_counts(now,kr_holiday):
    days = []
    for i in range(30):
        add_now = now + timedelta(days=i)
        if add_now.weekday() != 5 and add_now.weekday() != 6:
            if add_now not in kr_holiday:
                formattedDate = add_now.strftime("%#m/%#d")
                days.append(formattedDate)
    return days

# Get field in hwp file
def get_field_in_hwp(hwp,date,page_index):
    field_list = [i for i in hwp.GetFieldList().split('\x02')] # TODO 필드에 값 넣기
    for field in field_list:
        hwp.PutFieldText(f'{field}{{{{{page_index}}}}}',date)

def change_char_in_hwp(hwp): # now Use  / if you want change chars then you can use this funtion
    pass
    # all char search Aftor change funtion  
    # hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    # option=hwp.HParameterSet.HFindReplace
    # option.FindString = "오전"
    # option.ReplaceString = "3/27"
    # option.IgnoreMessage = 1
    # hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    # hwp.Clear(3)
    # hwp.Quit()

def get_now_year_holiday(now):
    # get kr holiday
    return [pytimekr.chuseok(now.year),pytimekr.lunar_newyear(now.year),pytimekr.hangul(now.year),pytimekr.children(now.year)
                ,pytimekr.independence(now.year),pytimekr.memorial(now.year),pytimekr.buddha(now.year)
                ,pytimekr.samiljeol(now.year),pytimekr.constitution(now.year)]

def ahto_print(path):
    # basic print auto
    rc = ShellExecuteEx(lpVerb = 'print',
                        lpFile = path,
                        fMask = shellcon.SEE_MASK_NOCLOSEPROCESS | shellcon.SEE_MASK_DOENVSUBST
                        # NOCLOSEPROCESS: 프로세스 핸들을 반환하도록 합니다.
                        # DOENVSUBST: lpFile에 포함된 환경 변수를 실제 값으로 바꿔주도록 합니다.
                        )
    hproc = rc['hProcess']
    win32event.WaitForSingleObject(hproc, win32event.INFINITE)
    exit_code = win32process.GetExitCodeProcess(hproc) 

# main
hwp = setting_hwp(path)
get_date(hwp,get_now_year_holiday(now),path,now)
#change_char_in_hwp(hwp)
