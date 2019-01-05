import subprocess, sys, os, re
from Tkinter import *

initial_window = Tk()
initial_window.resizable(width=False, height=False)
initial_window.title('compile tool')
initial_window.geometry('290x80')

pan_disk, patch_dts, select_version, module_type, top_second = 'D', '', '', '', ''

view_int, view_int_type = IntVar(), IntVar()
vvv, view_string_dts, view_string_serial = StringVar(), StringVar(), StringVar()
view_string_pan, view_string_button = StringVar(), StringVar()

SoftX3000_version = [
        ('R003C10', 0),
        ('R005C10', 1),
        ('R601C05', 2),
        ('R010C02', 3),
        ('R010C03', 4),
        ('R010C05', 5),
        ('R010C08', 6),
        ('R011C10', 7),
        ('R012C10', 8),
    ]

Module_Type = [
    ('fccu', 0),
    ('cdp', 1),
    ('ifm', 2),
    ('bsg', 3),
    ('msg', 4),
]

def get_cpci_compile_window():
    mytk = Toplevel()
    mytk.resizable(width=False, height=False)
    mytk.title("compile app")
    mytk.geometry("210x650")

    view_string_button.set('start compile')
    Label(mytk, text='[Note]', font=('black', 20), fg='green').pack()
    note_info = 'The default compilation path is [D:\] \nif change the path, write the new disk in the following ' \
                'path [E]'
    Label(mytk, text=note_info, font=('black', 10), fg='black', wraplength='210', justify='left').pack(anchor=W)

    Label(mytk, text='-- configure disk --', font=('black', 10), fg='black').pack()

    def set_configure_disk():
        global pan_disk, vvv, top_second
        pan_disk = vvv.get()
        with open('./configure.ini', 'w') as ff_file:
            tmp_pan = 'patch_disk=%s' % (pan_disk)
            ff_file.write(tmp_pan)
        top_second.destroy()

    def get_configure_disk():
        global pan_disk, vvv, top_second
        top_second = Toplevel()
        top_second.title('configure')
        e1 = Entry(top_second, textvariable=vvv, width=10)
        e1.grid(row=1, column=0, padx=1, pady=1)
        Button(top_second, text='set value', command=set_configure_disk).grid(row=1, column=1, padx=1, pady=1)

    def load_configure_disk():
        with open('./configure.ini', 'r') as ff_file:
            tmp_pan_disk = ff_file.read()
            disk = tmp_pan_disk.split('=')[1]
            return str(disk)

    Button(mytk, text='click configure disk', command=get_configure_disk).pack()
    Entry(mytk, text='')

    Label(mytk, text='-- select one version --', font=('black', 10), fg='black').pack(anchor=W)

    def get_version_choice():
        global select_version
        for index in range(12):
            if view_int.get() == index:
                select_version = SoftX3000_version[index][0]

    def get_module_type():
        global module_type
        for index in range(6):
            if view_int_type.get() == index:
                module_type = Module_Type[index][0]

    for lan, num in SoftX3000_version:
        Radiobutton(mytk, text=lan, value=num, command=get_version_choice, variable=view_int).pack(anchor=W)

    def get_patch_dts():
        global patch_dts
        patch_dts = view_string_dts.get()

    def get_serial_disk():
        global pan_disk
        if view_string_pan.get():
            pan_disk = view_string_pan.get()
        else:
            pan_disk = 'D'

    Label(mytk, text='-- patch_dts_codes --', font=('black', 10), fg='black').pack(anchor=W)
    Entry(mytk, textvariable=view_string_dts, width='27').pack(anchor=W)
    # Button(mytk, text='confirm path dts', command=get_patch_dts).pack(anchor=W)


    Label(mytk, text='-- select module type --', font=('black', 10), fg='black').pack(anchor=W)
    for lan, num in Module_Type:
        Radiobutton(mytk, text=lan, value=num, command=get_module_type, variable=view_int_type).pack(anchor=W)

    def start_compiling():
        global select_version, patch_dts, pan_disk, module_type, view_string_button
        view_string_button.set('compiled')
        get_patch_dts()
        get_module_type()
        pan_disk = load_configure_disk()
        command_path = r'%s:\OBJ\%s\build\makerom' % (pan_disk, select_version)
        cur_dirs = r'%s:\OBJ\%s\build\patch\%s' % (pan_disk, select_version, module_type)

        directory_int_set = set()
        for dir in os.listdir(cur_dirs):
            if re.search('patch', dir):
                directory_int_set.add(dir[5:9])

        max_path_num = int(max(directory_int_set)) + 1

        command_name = 'sx3kpatch.exe %s %s.c c %s' % (module_type, patch_dts, str(max_path_num))
        print command_path, command_name
        subprocess.check_call(command_name, shell=True, cwd=command_path)

    Button(mytk, text="start compiling", height='2', width='20', font=('black', 12), command=start_compiling,
           bg='#FFFAFA', fg='#4F4F4F', textvariable=view_string_button, activebackground='white', relief=RAISED).pack()

    mytk.mainloop()




def get_atca_compile_window():
    second_win = Toplevel()
    Label(second_win, text="lala").pack()

    putty_path = r'D:\HW iLMT\omu\workspace1\client\putty.exe'
    os.system(putty_path)
    compile_path = r'\\10.180.43.129\root\mnt\g00396313\build\makeatca'
    command_compile = r'./s3000make.sh clean acu'
    #subprocess.check_call(command_compile, shell=True, cwd=compile_path)
    
    # TODO #
    second_win.mainloop()



first_label = Label(initial_window, text='click compile CPCI', font=('black', 10), fg='black')
first_label.grid(row=0, column=0, sticky=W, padx=5,pady=5)
first_button = Button(initial_window, text='$ CPCI Window $', font=('black', 12), fg='black', command=get_cpci_compile_window)
first_button.grid(row=1, column=0, sticky=W, padx=5,pady=5)
first_label = Label(initial_window, text='click compile ATCA', font=('black', 10), fg='black')
first_label.grid(row=0, column=1, sticky=E, padx=5,pady=5)
second_button = Button(initial_window, text='& ATCA Window &', font=('black', 12), fg='black', command=get_atca_compile_window)
second_button.grid(row=1, column=1, sticky=W, padx=5, pady=5)

initial_window.mainloop()
