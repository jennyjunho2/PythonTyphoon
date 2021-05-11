import os
from time import sleep
from read_excel import array_to_problem_directory, get_problem_list

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
def find_word(hwp, word, direction="Forward"):
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir(direction)
    hwp.HParameterSet.HFindReplace.FindString = word
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HParameterSet.HFindReplace.SeveralWords = 1
    status = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    return

def source_to_basefile_copy_problem(hwp, source, problem_number):
    # 오류 처리 구문
    if os.path.exists(source) == False:
        raise Exception("문제저장용 파일이 존재하지 않습니다!")
    hwp.Open(rf'{source}')
    field_list = hwp.GetFieldList().split("\x02")
    if f'{problem_number}번문제' not in field_list:
        raise Exception(f"문제 저장용 파일에 {problem_number}번 문제가 존재하지 않습니다!")
    hwp.MoveToField(f'{problem_number}번문제')
    hwp.HAction.Run("MoveRight")
    start_pos = hwp.GetPos()
    end_string = """1번, 2번, 3번, 4번, 5번, 6번, 7번, 8번, 9번, 10번,
                 11번, 12번, 13번, 14번, 15번, 16번, 17번, 18번, 19번, 20번,
                 21번, 22번, 23번, 24번, 25번, 26번, 27번, 28번, 29번, 30번,
                 31번, 32번, 33번, 34번, 35번, 36번, 37번, 38번, 39번, 40번,
                 41번, 42번, 43번, 44번, 45번, 46번, 47번, 48번, 49번, 50번
                 """
    find_word(hwp, end_string, direction = "Forward")
    hwp.HAction.Run("MoveLineEnd")
    end_pos = hwp.GetPos()
    hwp.Run("Cancel")
    hwp.SetPos(*start_pos)
    hwp.Run("Select")
    hwp.SetPos(*end_pos)
    hwp.HAction.Run("Copy")
    hwp.HAction.Run("MoveDown")
    sleep(0.5)

def source_to_basefile_copy_solution(hwp, source, problem_number):
    if os.path.exists(source) == False:
        raise Exception("문제저장용 파일이 존재하지 않습니다!")
    hwp.Open(rf'{source}')
    field_list = hwp.GetFieldList().split("\x02")
    if f'{problem_number}번풀이' not in field_list:
        raise Exception(f"문제 저장용 파일에 {problem_number}번 풀이가 존재하지 않습니다!")
    hwp.MoveToField(f'{problem_number}번풀이')
    hwp.HAction.Run("MoveRight")
    hwp.HAction.Run("MoveSelTopLevelEnd")
    hwp.HAction.Run("Copy")
    hwp.HAction.Run("MoveDown")
    sleep(0.5)

"""
# 문제저장용 파일에서 검토용파일로 문제, 풀이를 붙여넣기하는 함수입니다.
# hwp : 아래아한글 기본파일
# destination : 검토용파일 파일경로
# problem_number : 검토용파일에 넣을 문제번호
# paste_problem은 문제를, paste_solution은 풀이를 복사하는 함수입니다.
# ★ 검토용파일에 누름틀(Ctrl + K E)이 '1번문제', '1번풀이' 양식으로 문제의 맨 앞에 지정되야 합니다!!!  
"""
def source_to_basefile_paste_problem(hwp, destination, problem_number):
    if os.path.exists(destination) == False:
        raise Exception("검토용 파일이 존재하지 않습니다!")
    hwp.Open(rf'{destination}')
    field_list = hwp.GetFieldList().split("\x02")
    if f'{problem_number}번문제' not in field_list:
        raise Exception(f"검토용 파일에 {problem_number}번 문제가 존재하지 않습니다!")
    hwp.MoveToField(f'{problem_number}번문제')
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("MoveDown")
    hwp.Save()
    sleep(0.5)

def source_to_basefile_paste_solution(hwp, destination, problem_number):
    if os.path.exists(destination) == False:
        raise Exception("검토용 파일이 존재하지 않습니다!")
    hwp.Open(rf'{destination}')
    field_list = hwp.GetFieldList().split("\x02")
    if f'{problem_number}번문제' not in field_list:
        raise Exception(f"검토용 파일에 {problem_number}번 풀이가 존재하지 않습니다!")
    hwp.MoveToField(f'{problem_number}번풀이')
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("MoveDown")
    hwp.Save()
    sleep(0.5)

# def usage_exclude(hwp, destination, grade):
#     print("출처 삭제중...")
#     hwp.Open(rf"{destination}")
#     hwp.MovePos(2)
#     field_list_problem = [field for field in hwp.GetFieldList().split("\x02") if "번문제" in field]
#     for field_problem in field_list_problem:
#         hwp.MoveToField(field_problem)
#         hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
#         if grade == 1:
#             hwp.HParameterSet.HFindReplace.FindString = "고1"
#         elif grade == 2:
#             hwp.HParameterSet.HFindReplace.FindString = "고2"
#         hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
#         hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
#         hwp.HParameterSet.HFindReplace.FindType = 1
#         status = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
#         hwp.HAction.Run("MoveNextParaBegin")
#         hwp.HAction.Run("MoveSelTopLevelEnd")
#         hwp.HAction.Run("Delete")
#         hwp.HAction.Run("MoveDown")
#     hwp.Save()
#     print("출처 삭제완료!")

def source_to_problem_execute(hwp, excel : str, grade_number : int, test_name : str,  dst : str):
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    for i in range(problems.shape[0]):
        src = array_to_problem_directory(problems[i, :], grade=grade_number)
        src_problem_number = get_problem_list(excel, grade=grade_number, test_name=test_name)[i][3]
        dst_problem_number = int(get_problem_list(excel, grade=grade_number, test_name=test_name)[i][4])
        # src_problem_score = int(get_problem_list(excel_test, grade=grade_number, test_name="테스트 시험지")[i][5])
        print(f"{dst_problem_number}번 입력중...({i+1}번째 입력)")
        source_to_basefile_problem(hwp, source = src , source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
        source_to_basefile_solution(hwp, source = src, source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
        print(f"{dst_problem_number}번 입력완료! ({i+1}번째 입력완료)")

"""
# 문제저장용 파일에서 베이스파일로 문제, 풀이를 복사, 붙여넣기하는 함수입니다.
# 위의 함수를 활용하였기에 copy, paste 함수만 잘 작동하면 됩니다!
# hwp : 아래아한글 기본파일
# source : 문제저장용 파일경로
# source_number : 문제저장용에서 가져올 문제 번호
# destination : 검토용파일 파일경로
# destination_number : 검토용파일에 넣을 문제번호  
"""
def source_to_basefile_problem(hwp, source, source_number, destination, destination_number):
    source_to_basefile_copy_problem(hwp, source, source_number)
    source_to_basefile_paste_problem(hwp, destination, destination_number)


def source_to_basefile_solution(hwp, source, source_number, destination, destination_number):
    source_to_basefile_copy_solution(hwp, source, source_number)
    source_to_basefile_paste_solution(hwp, destination, destination_number)
