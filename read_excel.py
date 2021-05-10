import os
import pandas as pd
from datetime import datetime
import directory_source

def readexcel(xlsx_path, grade : int):
    year = str(datetime.today().year)
    try:
        if grade == 1:
            xls_file = pd.read_excel(xlsx_path, sheet_name = '고1 '+ year)
        elif grade == 2:
            xls_file = pd.read_excel(xlsx_path, sheet_name = '고2 ' + year)
        else:
            raise Exception("학년을 다시 입력하여 주세요.")
        return xls_file
    except FileNotFoundError:
        print("주문서가 존재하지 않습니다!, 경로를 확인해 주세요.")

def get_problem_list(excel, grade : int, test_name : str):
    excel_problem_list = readexcel(excel, grade)
    score_index = excel_problem_list.columns.tolist().index(test_name)+1
    excel_problem_list_problem_index = excel_problem_list[['대단원', '소단원', '난이도','번호',test_name,
                                                           excel_problem_list.columns.tolist()[score_index]]].dropna(axis = 0).to_numpy()
    return excel_problem_list_problem_index

def array_to_problem_directory(array, grade : int):
    big_lesson_unit, small_lesson_unit, difficulty= array[0].replace(" ", ""), array[1].replace(" ", ""), array[2].replace(" ", "")
    if grade == 1:
        big_lesson = {'다항식' : 1,
                      '방정식' : 2,
                      '부등식' : 3,
                      '도형의방정식' : 4,
                      '집합과명제' : 5,
                      '함수' : 6,
                      '순열과조합' : 7}
        small_lesson = {'다항식의연산' : 1,
                        '항등식과나머지정리' : 2,
                        '인수분해' : 3,
                        '복소수' : 1,
                        '이차방정식' : 2,
                        '이차방정식과이차함수' : 3,
                        '여러가지방정식' : 4,
                        '일차부등식' : 1,
                        '이차부등식' : 2,
                        '평면좌표' : 1,
                        '직선의방정식' : 2,
                        '원의방정식' : 3,
                        '도형의이동' : 4,
                        '집합의뜻과표현' : 1,
                        '집합의연산' : 2,
                        '명제' : 3,
                        '함수' : 1,
                        '유리함수' : 2,
                        '무리함수' : 3,
                        '순열' : 1,
                        '조합' : 2}
        difficulty_number = {'최상' : 1,
                      '상' : 2,
                      '중' : 3,
                      '하' : 4}
    if grade == 2:
        big_lesson = {'지수함수와로그함수' : 1,
                      '삼각함수' : 2,
                      '수열' : 3,
                      '함수의극한과연속' : 4,
                      '미분' : 5,
                      '적분' : 6}
        small_lesson = {'지수' : 1,
                        '로그' : 2,
                        '지수함수' : 3,
                        '로그함수' : 4,
                        '삼각함수' : 1,
                        '삼각함수의그래프' : 2,
                        '삼각함수의활용' : 3,
                        '등차수열' : 1,
                        '등비수열' : 2,
                        '수열의합' : 3,
                        '수학적귀납법' : 4,
                        '함수의극한' : 1,
                        '함수의연속' : 2,
                        '미분계수와도함수' : 1,
                        '접선의방정식' : 2,
                        '함수의그래프' : 3,
                        '도함수의활용' : 4,
                        '부정적분' : 1,
                        '정적분' : 2,
                        '정적분의활용' : 3
                        }
        difficulty_number = {'최상': 1,
                      '상': 2,
                      '중': 3,
                      '하': 4}
    directory = str(grade)+"-"+str(big_lesson[big_lesson_unit])+'-'+str(small_lesson[small_lesson_unit])+'-'+str(difficulty_number[difficulty])
    directory_final = directory_source.directory_source[directory]
    return directory_final



if __name__ == "__main__":
    excelfile_directory = os.getcwd() + r'\태풍\내신주문서.xlsx'
    print(get_problem_list(excelfile_directory, 1, "테스트 시험지")[0])
    print(array_to_problem_directory(get_problem_list(excelfile_directory, 1, "테스트 시험지")[0], 1))
