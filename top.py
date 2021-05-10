import win32com.client as win32
from datetime import datetime as dt
from source_to_basefile import *
from basefile_to_exam import *
import os
from read_excel import *
import numpy as np

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.HAction.GetDefault("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    hwp.HParameterSet.HViewProperties.ZoomType = hwp.HwpZoomType("FitPage")
    hwp.HAction.Execute("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    return hwp

hwp = init_hwp()
directory = os.getcwd()
excel_test = directory + r'\태풍\내신주문서.xlsx'

grade_number = 1 # 고1
problems = get_problem_list(excel = excel_test, grade = grade_number, test_name = "테스트 시험지")
dst = directory + r"\태풍\testbench_dst.hwp"

# Test Code
start_time = dt.now()
for i in range(problems.shape[0]):
    src = array_to_problem_directory(problems[i, :], grade=grade_number)
    src_problem_number = get_problem_list(excel_test, grade=grade_number, test_name="테스트 시험지")[i][3]
    dst_problem_number = int(get_problem_list(excel_test, grade=grade_number, test_name="테스트 시험지")[i][4])
    # src_problem_score = int(get_problem_list(excel_test, grade=grade_number, test_name="테스트 시험지")[i][5])
    print(f"{dst_problem_number}번 입력중...({i+1}번째 입력)")
    source_to_basefile_problem(hwp, source = src , source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
    source_to_basefile_solution(hwp, source = src, source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
    print(f"{dst_problem_number}번 입력완료! ({i+1}번째 입력완료)")
print("출처 삭제중...")
usage_exclude(hwp, destination = dst)
print("출처 삭제완료!")
end_time = dt.now()
elapsed_time = end_time - start_time
print(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')