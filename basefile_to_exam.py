import shutil
import os
from time import sleep
import source_to_basefile

"""
# 배점 추가함수, 시험지에 문제를 추가한 후 실행하여주세요.
# 이 함수는 '?' 나 '하시오.' 위치를 기준으로 배점을 추가합니다.
# hwp : 아래아한글 실행
# source : 파일 위치
# problem_number : 문제 번호
# score : 추가할 배점
# choice : True면 객관식, False면 주관식 문제입니다.
# 아직 서술형 작은 문제 별 배점을 추가하진 않았습니다. 추후 업데이트 예정
"""

def insert_score(hwp, score):
    hwp.HAction.Run("MoveLineEnd")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = " "
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.GetDefault("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)
    hwp.HParameterSet.HEqEdit.Width = 2750
    hwp.HParameterSet.HEqEdit.EqFontName = "HYhwpEQ"
    hwp.HParameterSet.HEqEdit.string = f"left[{score}`점 right]"
    hwp.HParameterSet.HEqEdit.TreatAsChar = 1
    hwp.HAction.Execute("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)

def add_score(hwp, source, problem_number, score, choice = True):
    shutil.copyfile(source, rf"{source[:-4]}" + "_temp.hwp")
    source_temp = rf"{source[:-4]}" + "_temp.hwp"
    hwp.Open(f'{source_temp}')
    hwp.MoveToField(f'{problem_number}번문제')
    if choice == True:
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
        hwp.HParameterSet.HFindReplace.FindString = "?"
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        hwp.HParameterSet.HFindReplace.FindType = 1
        hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    elif choice == False:
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
        hwp.HParameterSet.HFindReplace.FindString = "하시오."
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        hwp.HParameterSet.HFindReplace.FindType = 1
        hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

    insert_score(hwp, score)
    hwp.SaveAs(rf"{source}")
    os.remove(source_temp)
    sleep(0.2)

def problem_array_to_exam_problem(hwp, array, grade : int, test_name : str, destination):
    for problem in array:
        directory_source, problem_number, exam_number, score = problem[0], problem[1], problem[2], problem[3]
        source_to_basefile.source_to_basefile_copy_problem(hwp = hwp, source = directory_source, problem_number = problem_number)
        source_to_basefile.source_to_basefile_paste_problem(hwp = hwp, destination = destination, problem_number = exam_number)

def main():
    pass

if __name__ == "__main__":
    main()