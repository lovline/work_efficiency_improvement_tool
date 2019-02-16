import tkinter as tk, os, time, sys

record_question = 'C:\Program Files (x86)\eSpace_Desktop\eSpace.exe'

questions_description, select_path = '', ''


# document path #
CPI_document = r'C:\Users\l00382665\Desktop\lovline_nana\HUAWEI UAC3000 接入网关控制器 产品文档-(V100R019C00_01).chm'
GPI_document = r'C:\Users\l00382665\Desktop\lovline_nana\HUAWEI UAC3000 接入网关控制器 GTS产品文档(GPI)-(V100R018C10_02).chm'
R011C10_document = r'C:\Users\l00382665\Desktop\lovline_nana\SoftX3000_CPI_disk_R11C10_cn.chm'


# execute program path #
chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
ie_path = r'C:\Program Files\Internet Explorer\iexplore.exe'
notepad_path = r'D:\Notepad++\notepad++.exe'
eDiary_path = r'D:\eDiary-3.3\eDiary.exe'
eSpace_path = r'C:\Program Files (x86)\eSpace_Desktop\eSpace.exe'
outlook_path = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013'
iLMT_path = r'D:\HW iLMT\omu\workspace1\client\lmt_client.exe'
source_insight_path = r'C:\Program Files (x86)\Source Insight 3\Insight3.exe'
RDO_path = r'C:\Program Files (x86)\Remote Desktop Organizer\RDO.exe'
pycharm_path = r'C:\Program Files\JetBrains\PyCharm 2017.1.5\bin\pycharm64.exe'

# directory path #
desktop_path = r'C:\Users\l00382665\Desktop'
lovline_path = r'C:\Users\l00382665\Desktop\lovline_nana'
work_path = r'N:\loveNN\2019下半年主要工作'
common_pathDDD = r'N:\CommonPatchDD'
work_national_path = r'N:\loveNN\2019下半年主要工作\National_Questions国内问题'
work_patch_path = r'N:\loveNN\2019下半年主要工作\paths补丁&需求处理'
work_national_path_month = r'N:\loveNN\2019下半年主要工作\National_Questions国内问题\2019年02月 爱卡西亚的暴雨'

open_execute_file_list = [
    CPI_document,
    GPI_document,
    R011C10_document,
    chrome_path,
    ie_path,
    notepad_path,
    eDiary_path,
    eSpace_path,
    outlook_path,
    iLMT_path,
    source_insight_path,
    source_insight_path,
    RDO_path,
    desktop_path,
    lovline_path,
    work_path,
    common_pathDDD,
    work_national_path,
    work_national_path,
    work_patch_path,
    work_national_path_month,
    # pycharm_path,
]


def start_execute():
    global open_execute_file_list
    for path in open_execute_file_list:
        os.startfile(path)
        time.sleep(0.7)
    window.quit()


window = tk.Tk()
window.title('One-click Open')
window.geometry('300x100')

tk.Label(window, textvariable='', width='27').pack()
tk.Button(window, text="open common program", height='2', width='20', font=('black', 12), command=start_execute,
           bg='#FFFAFA', fg='#4F4F4F', activebackground='white', relief='raised').pack()


window.mainloop()
