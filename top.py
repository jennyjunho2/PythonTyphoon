import win32com.client as win32
from datetime import datetime as dt
from source_to_basefile import *
from basefile_to_exam import *
import os
import shutil
from read_excel import *
from directory_source import *

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.HAction.GetDefault("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    hwp.HParameterSet.HViewProperties.ZoomType = hwp.HwpZoomType("FitPage")
    hwp.HAction.Execute("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    return hwp


hwp = init_hwp()
#############################
directory = os.getcwd()
problem_directory = directory + r'\문제저장용'
excel_test = directory + r'\태풍\내신주문서.xlsx'
testbench_fieldtest = directory + r'\태풍\testbench_fieldtest.hwp'
grade_number = 1 # 고1
#############################

# Test Code
# start_time = dt.now()
# # source_to_problem_execute(hwp, excel = excel_test, grade_number = grade_number, test_name = "테스트 시험지")
# source_to_problem_change_basefile(hwp, excel = excel_test, grade_number = grade_number, test_name_from = "테스트 시험지", test_name_to = "변형 시험지 1")
# end_time = dt.now()
# elapsed_time = end_time - start_time
# print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

#Test Code 2
# start_time = dt.now()
# add_field_source_file(hwp, testbench_fieldtest)
# end_time = dt.now()
# elapsed_time = end_time - start_time
# print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

print(get_problem_directory_list(problem_directory))