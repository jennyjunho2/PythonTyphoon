import win32com.client as win32
from datetime import datetime as dt
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
#############################

hwp.Open(r"D:\PythonTyphoon\태풍\testbench_fieldtest2.hwp")
hwp.MoveToField("1번문제글상자")
insert_text(hwp, "1번문제글상자 가능?")