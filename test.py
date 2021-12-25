import win32com.client as win32
import time
from source_to_basefile import *
from directory_source import *
import shutil

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = False
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

# hwp.Open(directory + r'\태풍\기출_문제+답지_1문제_test.hwp')
# max_number = 1
# for _ in range(600):
#     # field_list = hwp.GetFieldList().split("\x02")
#     hwp.MovePos(3)
#     insert_file(hwp, FileName=directory + r"\태풍\나를 붙여보세요.hwp")
#     hwp.RenameField("추가문제번호", f"{max_number + 1}번문제번호")
#     hwp.RenameField("추가풀이번호", f"{max_number + 1}번풀이번호")
#     hwp.RenameField("추가문제", f"{max_number + 1}번문제")
#     hwp.RenameField("추가풀이", f"{max_number + 1}번풀이")
#     hwp.PutFieldText(f"{max_number + 1}번문제번호", f"{max_number + 1}")
#     hwp.PutFieldText(f"{max_number + 1}번풀이번호", f"{max_number + 1}")
#     time.sleep(0.1)
#     max_number += 1

file_list = []
for path, subdirs, files in os.walk(r"C:\Users\Season\Desktop\기출시험지 배포용 - test\문제저장용"):
    for name in files:
        file_list.append(os.path.join(path, name))

for file_directory in file_list[99:]:
    print(file_directory, file_list.index(file_directory))
    hwp.Open(r"C:\Users\Season\Desktop\기출_문제+답지_1문제_test.hwp")
    destination = file_directory.replace("test", "destination")
    hwp.SaveAs(destination)
    hwp.Open(file_directory)
    problem_set = sorted(map(int, list(set(map(lambda x:x.replace("번문제", "").replace("번풀이", ""), get_field_list(hwp))))))
    for num in problem_set:
        source_to_basefile_problem(hwp, file_directory, num, destination, num, txtbox = False)
        source_to_basefile_solution(hwp, file_directory, num, destination, num, txtbox = False)
        hwp.Save()
# 고1 - 평면좌표 중(35)부터 해야함
# 고2는 지수 상(79)부터 시작
# 고2 - 지수 중(80)부터 진행해야함



