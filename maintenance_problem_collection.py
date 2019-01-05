import tkinter as tk, os, getpass, datetime, re, math
from docx import Document
import xlwt, xlrd, xlutils.copy

root_path = r'\\siarnd-fs\sia01\CNP_IMSCM_F\融合产品线维护部\1_NGN联合维护组项目文件夹\100 维护组\维护一组\37 现网保障&案例写作&问题记录'
record_question = '问题录入，格式如：【陕西电信】【1137689】入局呼叫甄别失败问题'

creator_chinese = {
    'l00238065': '李广',
    'h00396308': '何军',
    'l00474792': '李梦臣',
    'l00371553': '刘坚强',
    'l00382665': '刘伟',
    'w00190780': '王勇',
    'y00478622': '殷超超',
    'z00452218': '张伟',
    'z00381447': '周伟',
}

questions_type_directory = [
    ('00 AGCF&SIP类问题归总', 0),
    ('01 license类问题归总', 1),
    ('02 SOSM&ETSI&GB监听类问题归总', 2),
    ('03 SSF类问题归总', 3),
    ('04 uportal类问题归总', 4),
    ('05 补充业务问题归总', 5),
    ('06 彩铃传真类问题归总', 6),
    ('07 网管类问题归总', 7),
    ('08 过载类问题归总', 8),
    ('09 号码甄别&变换类问题归总', 9),
    ('10 话单类问题归总', 10),
    ('11 话统类问题归总', 11),
    ('12 扩容类问题归总', 12),
    ('13 前转类问题归总', 13),
    ('14 网关类问题归总', 14),
    ('15 消息跟踪类问题归总', 15),
    ('16 智能业务问题归总', 16),
    ('共性问题案例写作', 17),
]

def get_questions_description():
    return view_string_question.get()

def get_directory_choice():
    for index in range(len(questions_type_directory) + 1):
        if view_int_choice.get() == index:
            return questions_type_directory[index][0]

def write_question_to_excel(root_path, record_question, creator):
    record_excel_file_name = root_path + '\\' + '现网问题记录_录入的时候会自动填写.xls'
    product_information = 'SoftX3000'
    if re.match(r'SoftX3000|SX3000|SX|R010|R10|R011|R11', record_question, re.I):
        product_information = 'SoftX3000'
    elif re.match(r'UAC3000|UAC3.5|R003|UAC', record_question, re.I):
        product_information = 'UAC3000'
    elif re.match(r'uprotal', record_question, re.I):
        product_information = 'Uportal'
    region = '国内'
    is_public_flag = '否'
    question_state = 'OPEN'
    creator_str = creator_chinese[creator]
    icare_str = re.search(r'【\d{6,10}】|\d{6,10}', record_question)
    if icare_str is not None:
        icare_no = str(icare_str.group()).replace('【','').replace('】','')
    else:
        icare_no = ''
    site_information = record_question.split('】')[0].replace('【', '')
    if site_information == record_question:
        site_information = ''
    date_time = datetime.datetime.now().strftime('%Y-%m-%d')
    # 打开xls格式的excel文件 #
    excel_file = xlrd.open_workbook(filename=record_excel_file_name, formatting_info=True)
    table = excel_file.sheet_by_name('问题录入')
    # 得到当前行和列，新增数据要从nrow + 1行写入 #
    nrows = table.nrows
    ncol = table.ncols
    write_result_info = [nrows, date_time, product_information, region, site_information, record_question, is_public_flag,
                         creator_str, is_public_flag, question_state, icare_no, '', '']
    tmp_excel_file = xlutils.copy.copy(excel_file)
    tmp_table = tmp_excel_file.get_sheet(0)
    for col in range(ncol):
        tmp_table.write(nrows, col, write_result_info[col])
    tmp_excel_file.save(record_excel_file_name)

def start_create_and_open():
    # get user id #
    creator = getpass.getuser()
    select_path = get_directory_choice()
    record_question = get_questions_description()
    month_str = datetime.datetime.now().strftime('%Y-%m-%d')
    # month_str = now_time.replace('-', '_')
    if select_path == '共性问题案例写作':
        if record_question == '':
            # create a new document #
            document_name = 'XXX案例分享' + '_' + creator + '_' + month_str + '.docx'
        else:
            # create a description document #
            document_name = record_question + '_' + creator + '_' + month_str + '.docx'
        # 打开文档
        document = Document(docx=os.path.join(os.getcwd(), 'default.docx'))
        # document = Document()
        document.add_paragraph('')
        # 保存文件 #
        save_file_name = root_path + '\\' + select_path + '\\' + document_name
        document.save(save_file_name)
        os.startfile(root_path + '\\' + select_path)
        error_msg = '在--【%s】--时间作者--【%s】--写了一个问题案例：【%s】。' % (month_str, creator_chinese[creator], save_file_name)
        with open('log.txt', 'a') as log_file:
            log_file.writelines(error_msg)
    elif record_question is not '':
        # eg: 37 现网保障&案例写作&问题记录\09 号码甄别&变换类问题归总\【陕西电信】【SR 1137689】入局呼叫甄别失败问题_xxxxx_2019_xx_xx #
        result_path = root_path + '\\' + select_path + '\\' + record_question + '_' + creator + '_' + month_str
        print(result_path)
        # create the directory #
        os.makedirs(result_path)
        # open the directory #
        os.startfile(result_path)
        write_question_to_excel(root_path, record_question, creator)
        error_msg = '在--【%s】--时间作者--【%s】--创建了一个问题文件夹：【%s】。' % (month_str, creator_chinese[creator], result_path)
        with open('log.txt', 'a') as log_file:
            log_file.writelines(error_msg)
    else:
        error_msg = '非案例写作的情况下问题描述不能为空'
        with open('log.txt', 'a') as log_file:
            log_file.writelines(error_msg)
    window.quit()


window = tk.Tk()
window.title('Problem Collection Tool')
window.geometry('800x700')

view_int_choice, view_string_question = tk.IntVar() ,tk.StringVar()
tk.Label(window, text='-- 选择其中一个路径添加问题 --', bg='gray', font=('blue', 15), fg='black').pack(anchor='w')
for question, question_type in questions_type_directory:
    tk.Radiobutton(window, text=question, value=question_type, command=get_directory_choice, variable=view_int_choice).pack(anchor='w')

tk.Label(window, textvariable='', width='27').pack()
label = tk.Label(window,
    text=record_question,       # 标签的文字
    bg='gray',                 # 背景颜色
    font=('Arial', 10),         # 字体和字体大小
    width=100, height=2          # 标签长宽
    )
label.pack()    # 固定窗口位置
tk.Label(window, textvariable='', width='27').pack()
tk.Entry(window, textvariable=view_string_question, width=100).pack()
tk.Label(window, textvariable='', width='27').pack()

tk.Button(window, text="create and open", height='2', width='20', font=('black', 12), command=start_create_and_open,
           bg='#FFFAFA', fg='#4F4F4F', activebackground='white', relief='raised').pack()

window.mainloop()
