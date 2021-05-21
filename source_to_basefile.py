import os
import shutil
from time import sleep
from read_excel import *

"""
# 본 모듈은 문제저장용 파일에서 베이스파일로 가져올 때 사용하는 모듈입니다.
# 만든이 : 이준호(a01032208149@gmail.com으로 연락주세요.)
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

def write_text(hwp, text : str):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def source_to_basefile_copy_problem(hwp, source, problem_number):
    """
    # 문제저장용 파일에서 베이스파일로 문제, 풀이를 복사하는 함수입니다.
    # hwp : 아래아한글 기본파일
    # source : 문제저장용 파일경로
    # problem_number : 문제저장용 파일에서 가져올 문제번호
    # copy_problem은 문제를, copy_solution은 풀이를 복사하는 함수입니다.
    # ★ 문제저장용 파일에 누름틀(Ctrl + K E)이 '1번문제', '1번풀이' 양식으로 문제의 맨 앞에 지정되야 합니다!!!
    """
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
    sleep(0.2)

def add_problem_number_basefile(hwp, problem_array , file : str):
    hwp.Open(rf'{file}')
    field_list = [x for x in hwp.GetFieldList().split("\x02") if ('번호' in x)]
    field_list = [x for x in field_list if (int(x.replace("번풀이번호", "").replace("번문제번호", "")) < len(problem_array)+1)]
    for field in field_list:
        hwp.MoveToField(field)
        write_text(hwp, str(problem_array[int(field[0])-1]))

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
    sleep(0.2)

def source_to_basefile_paste_problem(hwp, destination, problem_number):
    """
    # 문제저장용 파일에서 검토용파일로 문제, 풀이를 붙여넣기하는 함수입니다.
    # hwp : 아래아한글 기본파일
    # destination : 검토용파일 파일경로
    # problem_number : 검토용파일에 넣을 문제번호
    # paste_problem은 문제를, paste_solution은 풀이를 복사하는 함수입니다.
    # ★ 검토용파일에 누름틀(Ctrl + K E)이 '1번문제', '1번풀이' 양식으로 문제의 맨 앞에 지정되야 합니다!!!
    """
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
    sleep(0.2)

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
    sleep(0.2)

def source_to_problem_execute(hwp, excel : str, grade_number : int, test_name : str):
    dst = new_basefile(test_name)
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    for i in range(problems.shape[0]):
        src = array_to_problem_directory(problems[i, :], grade=grade_number)
        src_problem_number = problems[i][3]
        dst_problem_number = int(problems[i][4])
        # src_problem_score = int(problems[i][5])
        print(f"{dst_problem_number}번 입력중...({i+1}번째 입력)")
        source_to_basefile_problem(hwp, source = src , source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
        source_to_basefile_solution(hwp, source = src, source_number = src_problem_number, destination = dst, destination_number = dst_problem_number)
        print(f"{dst_problem_number}번 입력완료! ({i+1}번째 입력완료)")

def source_to_problem_change_basefile(hwp, excel : str, grade_number : int, test_name_from : str, test_name_to : str):
    dst = new_basefile_no_number(test_name_to)
    problems_not_intersect = get_problem_list_change(excel = excel, grade = grade_number, test_name_from = test_name_from, test_name_to = test_name_to)[1]
    problem_number_list = []
    for i in range(problems_not_intersect.shape[0]):
        src = array_to_problem_directory(problems_not_intersect[i, :], grade=grade_number)
        src_problem_number = problems_not_intersect[i][3]
        dst_problem_number = i+1
        print(f"{dst_problem_number}번 입력중...({i+1}번째 입력)")
        source_to_basefile_problem(hwp, source=src, source_number=src_problem_number, destination=dst, destination_number=dst_problem_number)
        source_to_basefile_solution(hwp, source=src, source_number=src_problem_number, destination=dst, destination_number=dst_problem_number)
        print(f"{dst_problem_number}번 입력완료! ({i+1}번째 입력완료)")
        problem_number_list.append(int(problems_not_intersect[i][4]))
    problem_number_list =[replace_number_to_question[x] if int(x) > 40 else int(x) for x in problem_number_list]
    add_problem_number_basefile(hwp, problem_array = problem_number_list, file = dst)
    hwp.Save()

def source_to_basefile_problem(hwp, source, source_number, destination, destination_number):
    """
    # 문제저장용 파일에서 베이스파일로 문제를 복사, 붙여넣기하는 함수입니다.
    # 위의 함수를 활용하였기에 copy, paste 함수만 잘 작동하면 됩니다!
    # hwp : 아래아한글 기본파일
    # source : 문제저장용 파일경로
    # source_number : 문제저장용에서 가져올 문제 번호
    # destination : 검토용파일 파일경로
    # destination_number : 검토용파일에 넣을 문제번호
    """
    source_to_basefile_copy_problem(hwp, source, source_number)
    source_to_basefile_paste_problem(hwp, destination, destination_number)

def source_to_basefile_solution(hwp, source, source_number, destination, destination_number):
    """
       # 문제저장용 파일에서 베이스파일로 풀이를 복사, 붙여넣기하는 함수입니다.
       # 위의 함수를 활용하였기에 copy, paste 함수만 잘 작동하면 됩니다!
       # hwp : 아래아한글 기본파일
       # source : 문제저장용 파일경로
       # source_number : 문제저장용에서 가져올 문제 번호
       # destination : 검토용파일 파일경로
       # destination_number : 검토용파일에 넣을 문제번호
       """
    source_to_basefile_copy_solution(hwp, source, source_number)
    source_to_basefile_paste_solution(hwp, destination, destination_number)

def new_basefile(file_name : str):
    source_directory = os.getcwd() + r'\태풍\기출_문제+답지_원본_2문제씩_번호o.hwp'
    shutil.copyfile(source_directory, os.getcwd() + rf'\태풍\{file_name}_검토용파일(문제+답지).hwp')
    new_file = os.getcwd() + rf'\태풍\{file_name}_검토용파일(문제+답지).hwp'
    return new_file

def new_basefile_no_number(file_name : str):
    source_directory = os.getcwd() + r'\태풍\기출_문제+답지_원본_2문제씩_번호x.hwp'
    shutil.copyfile(source_directory, os.getcwd() + rf'\태풍\{file_name}_검토용파일(문제+답지).hwp')
    new_file = os.getcwd() + rf'\태풍\{file_name}_검토용파일(문제+답지).hwp'
    return new_file

def add_field(hwp, field_name):
    hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
    hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = field_name
    hwp.HParameterSet.HInsertFieldTemplate.TemplateName = field_name
    hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)


def add_field_source_file(hwp, source : str):
    hwp.Open(f"{source}")
    hwp.MovePos(2)
    for page in range(1, hwp.PageCount + 1):
        if page == 1:
            count = 0
            while count <= 7:
                hwp.HAction.Run("MoveDown")
                count = count + 1
            add_field(hwp, f"{page}번문제")
            hwp.MovePos(2)
            hwp.HAction.Run("MoveNextColumn")
            count2 = 0
            while count2 <= 7:
                hwp.HAction.Run("MoveDown")
                count2 = count2 + 1
            add_field(hwp, f"{page}번풀이")
            hwp.HAction.Run("MovePageDown")
            sleep(0.1)
        else:
            current_position = hwp.GetPos()
            hwp.HAction.Run("MovePrevPosEx")
            hwp.HAction.Run("MoveListBegin")
            add_field(hwp, f"{page}번문제")
            hwp.SetPos(*current_position)
            hwp.HAction.Run("MoveNextColumn")
            hwp.HAction.Run("MovePrevPosEx")
            hwp.HAction.Run("MoveListBegin")
            add_field(hwp, f"{page}번풀이")
            hwp.HAction.Run("MovePageDown")
            sleep(0.1)

def get_current_pos(hwp, source : str):
    hwp.Open(f"{source}")
    current_position = hwp.GetPos()
    return current_position

if __name__ == "__main__":
    excelfile_directory = os.getcwd() + r'\태풍\내신주문서.xlsx'