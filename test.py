import win32com.client as win32
import time
from source_to_basefile import *
from directory_source import *

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp

hwp = init_hwp()
#############################
directory = os.getcwd()
problem_directory = directory + r'\문제저장용'
problem_directory_test = directory + r'\문제저장용'
excel_test = directory + r'\태풍\내신주문서.xlsx'
excel_test_2 = directory+ r'\태풍\시험1.xlsx'
testbench_fieldtest = directory + r'\태풍\testbench_fieldtest.hwp'
testbench_fieldtest2 = directory + r'\태풍\testbench_fieldtest2.hwp'
testbench_fieldtest3 = directory + r'\태풍\testbench_fieldtest3.hwp'
testbench_basefiletest = directory + r'\태풍\testbench_basefiletest.hwp'
grade_number = 1 # 고1
###########################

hwp.Open(directory + r'\태풍\sample.hwp')
max_number = 260
for _ in range(50):
    # field_list = hwp.GetFieldList().split("\x02")
    hwp.MovePos(3)
    insert_file(hwp, FileName=directory + r"\태풍\나를 붙여보세요.hwp")
    hwp.RenameField("추가문제번호", f"{max_number + 1}번문제번호")
    hwp.RenameField("추가풀이번호", f"{max_number + 1}번풀이번호")
    hwp.RenameField("추가문제", f"{max_number + 1}번문제")
    hwp.RenameField("추가풀이", f"{max_number + 1}번풀이")
    hwp.PutFieldText(f"{max_number + 1}번문제번호", f"{max_number + 1}")
    hwp.PutFieldText(f"{max_number + 1}번풀이번호", f"{max_number + 1}")
    time.sleep(0.1)
    max_number += 1



