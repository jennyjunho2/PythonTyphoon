import win32com.client as win32
from datetime import datetime as dt
from source_to_basefile import *
from basefile_to_exam import *
import pandas as pd
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
problem_directory_test = directory + r'\문제저장용'
excel_test = directory + r'\태풍\내신주문서.xlsx'
excel_test_2 = directory+ r'\태풍\시험1.xlsx'
testbench_fieldtest = directory + r'\태풍\testbench_fieldtest.hwp'
testbench_fieldtest2 = directory + r'\태풍\testbench_fieldtest2.hwp'
testbench_fieldtest3 = directory + r'\태풍\testbench_fieldtest3.hwp'
testbench_basefiletest = directory + r'\태풍\testbench_basefiletest.hwp'
grade_number = 1 # 고1
#############################


# Test Code
# 엑셀을 통한 변형문제 제작
# start_time = dt.now()
# source_to_problem_execute(hwp, excel =r"C:\Users\Season\Desktop\준호타이핑용\testbench\태풍\시험1.xls", grade_number = grade_number, test_name = "1번 도형의 이동~집합과 명제 (210828)")
# source_to_problem_change_basefile(hwp, excel = excel_test, grade_number = grade_number, test_name_from = "테스트 시험지", test_name_to = "변형 시험지 1")
# end_time = dt.now()
# elapsed_time = end_time - start_time
# print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

# Test Code 2
# 'testbench_fieldtest'에 필드 추가하는 함수
# start_time = dt.now()
# add_field_source_file(hwp, r"D:\PythonTyphoon\태풍\3. 고2_수열_수열의 합_중.hwp")
# end_time = dt.now()
# elapsed_time = end_time - start_time
# print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

# Test Code 3
# 하위 경로의 모든 파일에 source file 추가
# start_time = dt.now()
# for file_directory in get_problem_directory_list(problem_directory_test):
#     add_field_source_file(hwp, source = file_directory)
# end_time = dt.now()
# elapsed_time = end_time - start_time
# print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

# Test Code 4
# 검토용파일 수정사항을 문제저장용에 반영
basefile_to_source(hwp = hwp, basefile=testbench_basefiletest, grade_number = grade_number, excel = excel_test_2)