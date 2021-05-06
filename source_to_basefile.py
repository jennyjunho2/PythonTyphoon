import shutil
import os
from time import sleep

"""
# 본 모듈은 문제저장용파일에서 베이스파일로 가져올 때 사용하는 모듈입니다.
# 만든이 : 이준호(a01032208149@gmail.com으로 연락주세요.)
"""

"""
# 문제저장용 파일에서 베이스파일로 문제, 풀이를 복사하는 함수입니다.
# hwp : 아래아한글 기본파일
# source : 문제저장용 파일경로
# problem_number : 문제저장용 파일에서 가져올 문제번호
# copy_problem은 문제를, copy_solution은 풀이를 복사하는 함수입니다.
# ★ 문제저장용 파일에 누름틀(Ctrl + K E)이 '1번문제', '1번풀이' 양식으로 문제의 맨 앞에 지정되야 합니다!!!
"""
def source_to_basefile_copy_problem(hwp, source, problem_number):
    hwp.Open(rf'{source}')
    hwp.MoveToField(f'{problem_number}번문제')
    hwp.HAction.Run("MoveRight")
    hwp.HAction.Run("MoveSelTopLevelEnd")
    hwp.HAction.Run("Copy")
    hwp.HAction.Run("MoveDown")
    sleep(0.5)

def source_to_basefile_copy_solution(hwp, source, problem_number):
    hwp.Open(rf'{source}')
    hwp.MoveToField(f'{problem_number}번풀이')
    hwp.HAction.Run("MoveRight")
    hwp.HAction.Run("MoveSelTopLevelEnd")
    hwp.HAction.Run("Copy")
    hwp.HAction.Run("MoveDown")
    sleep(0.5)

"""
# 문제저장용 파일에서 베이스파일로 문제, 풀이를 붙여넣기하는 함수입니다.
# hwp : 아래아한글 기본파일
# destination : 베이스파일 파일경로
# problem_number : 베이스파일에 넣을 문제번호
# paste_problem은 문제를, paste_solution은 풀이를 복사하는 함수입니다.
# ★ 베이스파일에 누름틀(Ctrl + K E)이 '1번문제', '1번풀이' 양식으로 문제의 맨 앞에 지정되야 합니다!!!  
"""
def source_to_basefile_paste_problem(hwp, destination, problem_number):
    shutil.copyfile(destination, rf"{destination[:-4]}"+"_temp.hwp")
    dst_temp = rf"{destination[:-4]}" + "_temp.hwp"
    hwp.Open(rf'{dst_temp}')
    hwp.MoveToField(f'{problem_number}번문제')
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("MoveDown")
    hwp.SaveAs(rf"{destination}")
    os.remove(dst_temp)
    sleep(0.5)

def source_to_basefile_paste_solution(hwp, destination, problem_number):
    shutil.copyfile(destination, rf"{destination[:-4]}"+"_temp.hwp")
    dst_temp = rf"{destination[:-4]}" + "_temp.hwp"
    hwp.Open(rf'{dst_temp}')
    hwp.MoveToField(f'{problem_number}번풀이')
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("MoveDown")
    hwp.SaveAs(rf"{destination}")
    os.remove(dst_temp)
    sleep(0.5)

"""
# 문제저장용 파일에서 베이스파일로 문제, 풀이를 복사, 붙여넣기하는 함수입니다.
# 위의 함수를 활용하였기에 copy, paste 함수만 잘 작동하면 됩니다!
# hwp : 아래아한글 기본파일
# source : 문제저장용 파일경로
# source_number : 문제저장용에서 가져올 문제 번호
# destination : 베이스파일 파일경로
# destination_number : 베이스파일에 넣을 문제번호  
"""
def source_to_basefile_problem(hwp, source, source_number, destination, destination_number):
    source_to_basefile_copy_problem(hwp, source, source_number)
    source_to_basefile_paste_problem(hwp, destination, destination_number)

def source_to_basefile_solution(hwp, source, source_number, destination, destination_number):
    source_to_basefile_copy_solution(hwp, source, source_number)
    source_to_basefile_paste_solution(hwp, destination, destination_number)
