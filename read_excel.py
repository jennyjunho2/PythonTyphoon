import os
import directory_source
import pandas as pd
from datetime import datetime
import warnings
import numpy as np
warnings.filterwarnings(action='ignore')

global replace_question_to_number
global replace_number_to_question
replace_question_to_number = { '단답형1' : 41,'단답형2' : 42, '단답형3' : 43,'단답형4' : 44,'단답형5' : 45,'단답형6' : 46,'단답형7' : 47,'단답형8' : 48,'단답형9' : 49,'단답형10' : 50,
                            '서술형1' : 51,'서술형2' : 52, '서술형3' : 53, '서술형4' : 54, '서술형5' : 55, '서술형6' : 56, '서술형7' : 57, '서술형8' : 58, '서술형9' : 59, '서술형10' : 60,
                               '주관식1' : 41, '주관식2' : 42, '주관식3' : 43, '주관식4' : 44, '주관식5' : 45, '주관식6' : 46, '주관식7' : 47,'주관식8' : 48,'주관식9' : 49,'주관식10' : 50}

replace_number_to_question = {41 : '단1', 42 : '단2', 43 : '단3',44 : '단4',45 : '단5',46 : '단6',47 : '단7', 48 : '단8', 49 : '단9', 50 : '단10',
                              51 : '서1', 52 : '서2', 53 : '서3',54 : '서4',55 : '서5',56 : '서6',57 : '서7', 58 : '서8', 59 : '서9', 60 : '서10'}

def readexcel(xlsx_path : str, grade : int):
    """
    readexcel 함수는 주문서 엑셀을 읽고 해당 학년의 sheet를 DataFrame의 형태로 반환합니다.
    :param xlsx_path: 엑셀 절대 경로
    :param grade: 해당 학년
    :return: DataFrame
    """
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

def get_problem_list(excel : str, grade : int, test_name : str):
    """
    get_problem_list는 주문서에서 해당 테스트에 해당하는 문제와 배점을 DataFrame 형태로 반환합니다.
    :param excel: 엑셀 절대 경로
    :param grade: 해당 학년
    :param test_name: 테스트 이름
    :return: DataFrame
    """
    excel_problem_list = readexcel(excel, grade)
    excel_problem_list_problem_index = excel_problem_list[['대단원', '소단원', '난이도','번호',test_name,
                                                           excel_problem_list.columns.tolist()[excel_problem_list.columns.tolist().index(test_name)+1]]].dropna(axis = 0).replace({test_name : replace_question_to_number})
    return excel_problem_list_problem_index.sort_values(by = [test_name])

def array_to_problem_directory(problem_set, grade : int, test_name : str, for_release : bool, ip_address = "172.30.1.60", test = False):
    """
    대단원, 소단원, 난이도, 번호, 테스트 이름, 배점의 형태를 읽어 이를 hwp 경로로 읽을 수 있게 변환하여주는 함수입니다.\n
    :param problem_set: array
    :param grade: 해당 학년
    :param test_name: 테스트 이름
    :param for_release: 배포용에서 상호작용 할 것인지?
    :return: string(ex. 1-1-1-1)
    """
    directory_final = []
    problem_set.index = ["대단원", '소단원', '난이도', '번호', test_name, '배점']
    big_lesson_unit, small_lesson_unit, difficulty= problem_set['대단원'].replace(" ", ""), problem_set['소단원'].replace(" ", ""), problem_set['난이도'].replace(" ", "")
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
    if test == False:
        if for_release == True:
            directory_final.extend((directory_source.directory_source(directory, r"C:\Users\Season\Desktop\기출시험지 배포용"), problem_set['번호'], problem_set[test_name], problem_set['배점']))
        else:
            directory_final.extend((directory_source.directory_source(directory, rf"\\{ip_address}\기출시험지"), problem_set['번호'], problem_set[test_name], problem_set['배점']))
    else:
        directory_final.extend((directory_source.directory_source(directory, r"D:\PythonTyphoon"), problem_set['번호'], problem_set[test_name], problem_set['배점']))
    return directory_final

def get_problem_list_change(excel, grade : int, test_name_from : str, test_name_to : str):
    """
    엑셀 DataFrame에서 두 테스트 간 문제 배열을 비교하여 겹치는 부분과 겹치지 않는 부분을 반환합니다.
    :param excel: 엑셀 절대 경로
    :param grade: 해당 학년
    :param test_name_from: 비교할 시험지 이름
    :param test_name_to: 비교 및 제작할 시험지 이름
    :return: 겹치는 부분 array, 겹치지 않는 부분 array
    """
    excel_problem_list = readexcel(excel, grade)
    excel_problem_list_problem_index = excel_problem_list[['대단원', '소단원', '난이도','번호',test_name_from,
                                                           excel_problem_list.columns.tolist()[excel_problem_list.columns.tolist().index(test_name_from) + 1],
                                                           test_name_to, excel_problem_list.columns.tolist()[excel_problem_list.columns.tolist().index(test_name_to) + 1]]].replace({test_name_to : replace_question_to_number}).replace({test_name_from : replace_question_to_number})
    excel_problem_list_intersect = excel_problem_list_problem_index.dropna(axis = 0).iloc[:, [0, 1, 2, 3, 6, 7]].sort_values(by = [test_name_to], axis = 0).to_numpy()
    excel_problem_list_not_intersect = excel_problem_list_problem_index.dropna(subset = [test_name_to])[excel_problem_list_problem_index[test_name_from].isna()].iloc[:, [0, 1, 2, 3, 6, 7]].sort_values(by = [test_name_to], axis = 0).to_numpy()
    return excel_problem_list_intersect, excel_problem_list_not_intersect

if __name__ == "__main__":
    excelfile_directory = os.getcwd() + r'\태풍\내신주문서.xlsx'
    # print(readexcel(excelfile_directory, 1))
    # print(get_problem_list(excel = r"D:\PythonTyphoon\태풍\내신주문서.xlsx", grade = 1, test_name = "양정 최종마무리3 (210707)"))