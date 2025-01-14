import sys
import win32com.client as win32
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon, QKeySequence
from source_to_basefile import *
from api import *
import shutil
import time
import re

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp

hwp = init_hwp()
#######################################################################################################################
class WindowClass(QDialog):
    first_click = True

    def __init__(self):
        
        super().__init__()
        os.chdir(r"C:\Users\SEASON\Desktop\ljh\testbench\PythonTyphoon")
        self.setWindowIcon(QIcon(os.getcwd() + r"\typhoon_gui\icon.png"))
        self.ui = uic.loadUi(os.getcwd() + r"\typhoon_gui\test.ui", self)
        self.setMaximumWidth(self.width())
        self.setMinimumWidth(self.width())
        self.setMaximumHeight(self.height())
        self.setMinimumWidth(self.width())

        self.setWindowTitle("검토용파일 제작 프로그램")
        self.ui.closeEvent = self.closeEvent

        self.pushButton_execute.clicked.connect(self.execute_function)
        self.pushButton_execute_2.clicked.connect(self.execute_function_2)
        self.pushButton_execute_3.clicked.connect(self.execute_function_3)
        self.pushButton_find_excel.clicked.connect(self.get_excel_file_name)
        self.pushButton_find_excel_2.clicked.connect(self.get_excel_file_name_2)
        self.pushButton_find_hwp_2.clicked.connect(self.get_hwp_file_name_2)
        self.pushButton_find_excel_3.clicked.connect(self.get_excel_file_name_3)
        self.pushButton_change_ip_address.clicked.connect(self.change_ip_address)

        self.shortcut_excel = QShortcut(QKeySequence("Alt+E"), self)
        self.shortcut_hwp = QShortcut(QKeySequence("Alt+H"), self)
        self.shortcut_excel.activated.connect(self.get_excel_file_name)
        self.tab_kind.tabBarClicked.connect(self.handle_tab_bar_clicked)

        self.shortcut_tab_first = QShortcut(QKeySequence("Ctrl+1"), self)
        self.shortcut_tab_second = QShortcut(QKeySequence("Ctrl+2"), self)
        self.shortcut_tab_third = QShortcut(QKeySequence("Ctrl+3"), self)
        self.shortcut_tab_first.activated.connect(self.change_tab_to_first)
        self.shortcut_tab_second.activated.connect(self.change_tab_to_second)
        self.shortcut_tab_third.activated.connect(self.change_tab_to_third)

        #self.shortcut_execute_1 = QShortcut(QKeySequence("Ctrl+E"), self)
        #self.shortcut_execute_2 = QShortcut(QKeySequence("Ctrl+E"), self)
        #self.shortcut_execute_3 = QShortcut(QKeySequence("Ctrl+E"), self)

        # self.shorcut_execute_1.activated.connect(self.execute_function)
        #self.shortcut_execute_2.activated.connect(self.execute_function_2)
        #self.shortcut_execute_3.activated.connect(self.execute_function_3)


        self.get_ip_address()

    def change_tab_to_first(self):
        if self.tab_kind.currentIndex != 0:
            self.tab_kind.setCurrentIndex(0)
            try: self.shortcut_excel.activated.disconnect()
            except TypeError: pass
            try: self.shortcut_hwp.activated.disconnect()
            except TypeError: pass

            self.shortcut_excel.activated.connect(self.get_excel_file_name)
        else:
            pass

    def change_tab_to_second(self):
        if self.tab_kind.currentIndex != 1:
            self.tab_kind.setCurrentIndex(1)
            try: self.shortcut_excel.activated.disconnect()
            except TypeError: pass
            try: self.shortcut_hwp.activated.disconnect()
            except TypeError: pass

            self.shortcut_excel.activated.connect(self.get_excel_file_name_2)
            self.shortcut_hwp.activated.connect(self.get_hwp_file_name_2)
        else:
            pass

    def change_tab_to_third(self):
        if self.tab_kind.currentIndex != 2:
            self.tab_kind.setCurrentIndex(2)
            try: self.shortcut_excel.activated.disconnect()
            except TypeError: pass
            try: self.shortcut_hwp.activated.disconnect()
            except TypeError: pass

            self.shortcut_excel.activated.connect(self.get_excel_file_name_3)
        else:
            pass

    def handle_tab_bar_clicked(self, index):
        try: self.shortcut_excel.activated.disconnect()
        except TypeError: pass
        try: self.shortcut_hwp.activated.disconnect()
        except TypeError: pass

        if index == 0:
            self.shortcut_excel.activated.connect(self.get_excel_file_name)
        elif index == 1:
            self.shortcut_excel.activated.connect(self.get_excel_file_name_2)
            self.shortcut_hwp.activated.connect(self.get_hwp_file_name_2)
        elif index == 2:
            self.shortcut_excel.activated.connect(self.get_excel_file_name_3)

    def get_ip_address(self):
        with open(os.getcwd() + r"\ip_address.txt", 'r') as ip_address_file:
            ip_address = ip_address_file.readline()
        self.QTextEdit_ip_address.setPlainText(ip_address)

    def change_ip_address(self):
        new_ip = self.QTextEdit_ip_address.toPlainText()
        with open(os.getcwd() + r"\ip_address.txt", 'w') as ip_address_file:
            ip_address_file.write(new_ip)
        myWindow.append_text_function("IP 주소가 업데이트 되었습니다!")

    def closeEvent(self, event):
        hwp.Quit()

    def append_text_function(self, string):
        self.textBrowser_progress.append(rf"{string}")
        QCoreApplication.processEvents()

    def get_excel_file_name(self):
        file_filter = 'Excel File (*.xlsx *.xls)'
        response = QFileDialog.getSaveFileName(
            parent=self,
            caption="엑셀 파일을 선택해주세요.",
            filter=file_filter,
            initialFilter="엑셀 파일 (*.xlsx *.xls)",
            directory=os.getcwd()
        )
        self.QTextEdit_excel_directory.setPlainText(response[0])

    def execute_function(self):
        if self.radioButton_grade1.isChecked():
            grade_number = 1
        elif self.radioButton_grade2.isChecked():
            grade_number = 2
        if self.checkBox_testmode.isChecked():
            test_mode = True
        else:
            test_mode = False
        excel_directory = self.QTextEdit_excel_directory.toPlainText().strip('""')
        test_name = self.QTextEdit_test_name.toPlainText().strip('""')
        ip_address = self.QTextEdit_ip_address.toPlainText().strip('""')
        try:
            sleep(2)
            self.pushButton_execute.setEnabled(False)
            hwp = init_hwp()
            source_to_problem_execute_gui(hwp=hwp, excel=excel_directory, test_name=test_name,
                                          grade_number=grade_number, ip_address=ip_address, test_mode=test_mode)
            self.pushButton_execute.setEnabled(True)
        except OSError as e:
            myWindow.append_text_function(string="주문서 경로를 확인해주세요. : " + str(e))
            self.pushButton_execute.setEnabled(True)
        except UnboundLocalError as e:
            myWindow.append_text_function(string="학년을 선택해주세요. : " + str(e))
            self.pushButton_execute.setEnabled(True)
        except ValueError as e:
            myWindow.append_text_function(string="학년과 시험지 이름을 확인해주세요. : " + str(e))
            self.pushButton_execute.setEnabled(True)
        except Exception as e:
            myWindow.append_text_function(string="오류가 발생했습니다 : " + str(e))
            self.pushButton_execute.setEnabled(True)

    ################################################################################################################

    def append_text_function_2(self, string):
        self.textBrowser_progress_2.append(rf"{string}")
        QCoreApplication.processEvents()

    def get_excel_file_name_2(self):
        file_filter = 'Excel File (*.xlsx *.xls)'
        response = QFileDialog.getSaveFileName(
            parent=self,
            caption="엑셀 파일을 선택해주세요.",
            filter=file_filter,
            initialFilter="엑셀 파일 (*.xlsx *.xls)",
            directory=os.getcwd()
        )
        self.QTextEdit_excel_directory_2.setPlainText(response[0])

    def get_hwp_file_name_2(self):
        file_filter = 'Hwp File (*.hwp)'
        response = QFileDialog.getSaveFileName(
            parent=self,
            caption="한글 파일을 선택해주세요.",
            filter=file_filter,
            initialFilter="한글 파일 (*.hwp)",
            directory=r"\\Desktop-r3553ul\기출시험지\문제저장용\문제+답지"
        )
        self.QTextEdit_hwp_directory_2.setPlainText(response[0])

    def execute_function_2(self):
        if self.radioButton_grade1_2.isChecked():
            grade_number = 1
        elif self.radioButton_grade2_2.isChecked():
            grade_number = 2
        if self.checkBox_testmode.isChecked():
            test_mode = True
        else:
            test_mode = False
        excel_directory = self.QTextEdit_excel_directory_2.toPlainText().strip('""')
        basefile_directory = self.QTextEdit_hwp_directory_2.toPlainText().strip('""')
        test_name = self.QTextEdit_test_name_2.toPlainText().strip('""')
        ip_address = self.QTextEdit_ip_address.toPlainText().strip('""')
        try:
            self.pushButton_execute_2.setEnabled(False)
            basefile_to_source_gui(hwp=hwp, excel=excel_directory, basefile=basefile_directory, test_name=test_name,
                                   grade_number=grade_number, ip_address=ip_address, test_mode=test_mode)
            self.pushButton_execute_2.setEnabled(True)
        except OSError:
            myWindow.append_text_function_2(string="검토용파일 경로를 확인해주세요.")
            self.pushButton_execute_2.setEnabled(True)
        except UnboundLocalError:
            myWindow.append_text_function_2(string="학년을 선택해주세요.")
            self.pushButton_execute_2.setEnabled(True)
        except ValueError:
            myWindow.append_text_function_2(string="학년과 시험지 이름을 확인해주세요.")
            self.pushButton_execute_2.setEnabled(True)
        except Exception as e:
            myWindow.append_text_function_2(string="오류가 발생했습니다 : " + str(e))
            self.pushButton_execute_2.setEnabled(True)

    ################################################################################################################

    def append_text_function_3(self, string):
        self.textBrowser_progress_3.append(rf"{string}")
        QCoreApplication.processEvents()

    def get_excel_file_name_3(self):
        file_filter = 'Excel File (*.xlsx *.xls)'
        response = QFileDialog.getSaveFileName(
            parent=self,
            caption="엑셀 파일을 선택해주세요.",
            filter=file_filter,
            initialFilter="엑셀 파일 (*.xlsx *.xls)",
            directory=os.getcwd()
        )
        self.QTextEdit_excel_directory_3.setPlainText(response[0])

    def execute_function_3(self):
        if self.radioButton_grade1_3.isChecked():
            grade_number = 1
        elif self.radioButton_grade2_3.isChecked():
            grade_number = 2
        if self.checkBox_testmode.isChecked():
            test_mode = True
        else:
            test_mode = False
        excel_directory = self.QTextEdit_excel_directory_3.toPlainText().strip('""')
        test_name = self.QTextEdit_test_name_3.toPlainText().strip('""')
        ip_address = self.QTextEdit_ip_address.toPlainText().strip('""')
        reference_string = self.QTextEdit_reference_string_3.toPlainText().strip('""')
        try:
            self.pushButton_execute_3.setEnabled(False)
            add_reference_gui(hwp=hwp, test_name=test_name, grade_number=grade_number, ip_address=ip_address,
                              reference_string=reference_string, excel=excel_directory, test_mode=test_mode)
            self.pushButton_execute_3.setEnabled(True)
        except OSError:
            myWindow.append_text_function_3(string="검토용파일 경로를 확인해주세요.")
            self.pushButton_execute_3.setEnabled(True)
        except UnboundLocalError:
            myWindow.append_text_function_3(string="학년을 선택해주세요.")
            self.pushButton_execute_3.setEnabled(True)
        except ValueError:
            myWindow.append_text_function_3(string="학년과 시험지 이름을 확인해주세요.")
            self.pushButton_execute_3.setEnabled(True)
        except Exception as e:
            myWindow.append_text_function_3(string="오류가 발생했습니다 : " + str(e))
            self.pushButton_execute_3.setEnabled(True)


def new_basefile_no_number_gui(file_name: str, grade_number, test_mode=False):
    if not test_mode:
        source_directory = r"C:\Users\Season\Desktop\자동화\기출_문제+답지_원본_2문제씩_번호x.hwp"
        parent_directory = r"C:\Users\Season\Desktop\자동화\\"
    else:
        source_directory = r"D:\PythonTyphoon\태풍\기출_문제+답지_원본_2문제씩_번호x.hwp"
        parent_directory = r"D:\PythonTyphoon\태풍\\"

    if "최종마무리" not in file_name:
        try:
            file_name_date = re.search(r"\(([0-9_]+)\)", file_name).group(1)
        except AttributeError:
            file_name_date = ""
        try:
            file_name_school = re.search(r"\(([가-힣^,]+)\)", file_name).group(1)
        except AttributeError:
            file_name_school = ""
        file_name_count = file_name[:file_name.find("번")] if '번' in file_name else 0
        file_name = file_name.replace(f"({file_name_date})", "").replace(f"{file_name_count}번", "").replace(
            f"{file_name_school}", "").replace("()", "")
        if file_name_count != 0:
            file_copy_directory = parent_directory + str(file_name_date) + f"_고" + str(grade_number) + " " + str(
                file_name_count) + "_" + str(file_name_school) + "(" + file_name[1:].strip() + str(")_검토용파일.hwp")
        else:
            file_copy_directory = parent_directory + str(file_name_date) + f"_고" + str(grade_number) + "_" + str(
                file_name_school) + "(" + file_name[1:].strip() + str(")_검토용파일.hwp")
    else:
        try:
            file_name_date = re.search(r"\(([0-9_]+)\)", file_name).group(1)
        except AttributeError:
            file_name_date = ""
        try:
            file_name_school = re.search(r"[가-힣]{2,3}", file_name).group()
        except AttributeError:
            file_name_school = ""
        file_name = file_name.replace(f"({file_name_date})", "").replace(f"{file_name_school}", "").replace("()", "")
        file_copy_directory = parent_directory + str(file_name_date) + f"_고" + str(grade_number) + "_" + str(
            file_name_school) + " " + file_name.strip() + str("_검토용파일.hwp")
    shutil.copyfile(source_directory, file_copy_directory)
    new_file = file_copy_directory
    return new_file


def source_to_problem_execute_gui(hwp, excel: str, grade_number: int, test_name: str, ip_address: str,
                                  basefile: bool = True, test_mode: bool = False):
    start_time = dt.now()
    if not test_mode:
        myWindow.append_text_function(string=f"{test_name} 제작 시작")
    else:
        myWindow.append_text_function(string=f"{test_name} 제작 시작 (테스트모드)")

    progress = 0
    myWindow.progressbar.setValue(progress)
    myWindow.append_text_function(string="엑셀 로딩중...")
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    myWindow.append_text_function(string="엑셀 로딩 완료")

    dst_problem_number_for_field = [x for x in range(1, len(readexcel(excel, grade=grade_number)[test_name]) + 1)]
    dst = new_basefile_no_number_gui(test_name, grade_number=grade_number, test_mode=test_mode)
    for i in range(problems.shape[0]):
        myWindow.append_text_function(string=f"{dst_problem_number_for_field[i]}번 입력중...({i + 1}번째 입력)")
        problem_set = problems.iloc[i]
        src = array_to_problem_directory(problem_set, grade=grade_number, test_name=test_name, for_release=True,
                                         test=test_mode, ip_address=ip_address)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        source_to_basefile_problem(hwp, source=problem_directory, source_number=src_problem_number, destination=dst,
                                   destination_number=dst_problem_number_for_field[i])
        source_to_basefile_solution(hwp, source=problem_directory, source_number=src_problem_number, destination=dst,
                                    destination_number=dst_problem_number_for_field[i])
        hwp.PutFieldText(Field=f"{i + 1}번문제번호", Text=str(replace_number_to_question[int(dst_problem_number)]) if int(
            dst_problem_number) >= 41 else str(int(dst_problem_number)))
        hwp.PutFieldText(Field=f"{i + 1}번풀이번호", Text=str(replace_number_to_question[int(dst_problem_number)]) if int(
            dst_problem_number) >= 41 else str(int(dst_problem_number)))
        myWindow.append_text_function(string=f"{dst_problem_number_for_field[i]}번 입력완료! ({i + 1}번째 입력완료)")
        hwp.Save()
        progress += (100 // problems.shape[0])
        myWindow.progressbar.setValue(progress)
    if basefile:
        hwp.PutFieldText(Field="검토용파일이름", Text=test_name)

    # 남은 페이지 삭제하기
    needed_page = int(np.ceil(problems.shape[0] / 2))
    hwp.MovePos(2)
    for _ in range(needed_page):
        hwp.HAction.Run("MovePageDown")
    hwp.MovePos(7)
    start_pos = hwp.GetPos()
    hwp.MovePos(3)
    end_pos = hwp.GetPos()
    hwp.SetPos(*start_pos)
    hwp.Run("Select")
    hwp.SetPos(*end_pos)
    hwp.HAction.Run("Delete")

    # 제목 넣기
    hwp.MoveToField("검토용파일제목")
    if '최종마무리' not in test_name:
        try:
            file_name_school = re.search(r"\(([가-힣^,]+)\)", test_name).group(1)
        except AttributeError:
            file_name_school = ""
        file_name_count = test_name[:test_name.find("번")] if '번' in test_name else 0
        insert_text(hwp, string=f"{str(grade_number)} ")
        circle_word(hwp, string=f"{str(file_name_count)}")
        insert_text(hwp, string=f" {str(file_name_school)}")
    else:
        try:
            file_name_school = re.search(r"[가-힣]{2,3}", test_name).group()
        except AttributeError:
            file_name_school = ""
        insert_text(hwp, string=f"{str(grade_number)} {file_name_school} 최종마무리 ")
        circle_word(hwp, string=f"{test_name[test_name.find('리') + 1]}")

    # 누름틀 전체 삭제
    # new_field_list = hwp.GetFieldList().split("\x02")
    # for field in new_field_list:
    #     hwp.MoveToField(f"{field}")
    #     hwp.HAction.Run("DeleteField")

    hwp.Save()
    sleep(0.2)
    end_time = dt.now()
    elapsed_time = end_time - start_time
    myWindow.append_text_function(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')
    myWindow.progressbar.setValue(100)
    hwp.Quit()


def basefile_to_source_gui(hwp, basefile: str, test_name: str, grade_number, ip_address: str, excel=None,
                           test_mode: bool = False):
    end_string = """1번, 2번, 3번, 4번, 5번, 6번, 7번, 8번, 9번, 10번,
                     11번, 12번, 13번, 14번, 15번, 16번, 17번, 18번, 19번, 20번,
                     21번, 22번, 23번, 24번, 25번, 26번, 27번, 28번, 29번, 30번,
                     31번, 32번, 33번, 34번, 35번, 36번, 37번, 38번, 39번, 40번,
                     41번, 42번, 43번, 44번, 45번, 46번, 47번, 48번, 49번, 50번
                     """
    end_string2 = """풀이, 타이핑"""
    start_time = dt.now()
    hwp.Open(rf'{basefile}')
    # 검토용파일 존재하는지 검사
    if not os.path.exists(basefile):
        raise Exception("검토용파일이 존재하지 않습니다!")
    if not test_mode:
        myWindow.append_text_function_2(string=f"{basefile} 반영 시작")
    else:
        myWindow.append_text_function_2(string=f"{basefile} 반영 시작 (테스트모드)")
    progress = 0
    myWindow.progressbar_2.setValue(progress)
    myWindow.append_text_function_2(string=test_name + " 반영 진행중...")
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    field_list = hwp.GetFieldList().split("\x02")
    field_list_problem_number = [x for x in field_list if "번문제번호" in x]
    field_list_solution_number = [x for x in field_list if "번풀이번호" in x]
    field_list_change_problem_number = []
    field_list_change_solution_number = []
    # 문제 번호 빨간색 표시 검사
    for field_problem_number in field_list_problem_number:
        hwp.MoveToField(field_problem_number, start=False)
        hwp.HAction.Run("SelectAll")
        color = hex(hwp.CharShape.Item("TextColor"))
        
        if hwp.CharShape.Item("TextColor") == hex(255):  # 빨간색일 경우
            hwp.HAction.Run("MoveLeft")
            field_list_change_problem_number.append(hwp.GetCurFieldName())
            hwp.HAction.Run("SelectAll")
            # hwp.HAction.Run("CharShapeTextColorBlack")
        else:
            pass

    for field_solution_number in field_list_solution_number:
        hwp.MoveToField(field_solution_number, start=False)
        hwp.HAction.Run("SelectAll")
        if hwp.CharShape.Item("TextColor") == hex(255):  # 빨간색일 경우
            hwp.HAction.Run("MoveLeft")
            field_list_change_solution_number.append(hwp.GetCurFieldName())
            hwp.HAction.Run("SelectAll")
            # hwp.HAction.Run("CharShapeTextColorBlack")
        else:
            pass

    field_list_change_problem_number = list(map(lambda y: int(y) - 1, list(map(lambda x: x[:-5], field_list_change_problem_number))))
    field_list_change_solution_number = list(map(lambda y: int(y) - 1, list(map(lambda x: x[:-5], field_list_change_solution_number))))
    problem_change_problem = problems.iloc[field_list_change_problem_number]
    problem_change_solution = problems.iloc[field_list_change_solution_number]
    hwp.Save()

    problem_read_only = []
    solution_read_only = []
    for i in range(problem_change_problem.shape[0]):  # 각 문제에 대하여
        problem_change_problem_set = problem_change_problem.iloc[i]
        src = array_to_problem_directory(problem_change_problem_set, grade=grade_number, test_name=test_name,
                                         for_release=False, ip_address=ip_address, test=test_mode)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        myWindow.append_text_function_2(string=f"{field_list_change_problem_number[i] + 1}번문제 반영중...({i + 1}번째 입력)")

        # 반영 전 검토용파일 문제 복사하기
        hwp.Open(rf'{basefile}')
        if (hwp.EditMode == 0):
            problem_read_only.append(dst_problem_number)
            continue

        hwp.MoveToField(f"{field_list_change_problem_number[i] + 1}번문제글상자", start=True)
        start_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.MoveToField(f"{field_list_change_problem_number[i] + 1}번문제글상자", start=False)
        end_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.HAction.Run("Copy")
        hwp.HAction.Run("MoveDown")

        # 문제저장용에 붙여넣기
        if not os.path.exists(problem_directory):
            raise Exception(f"{field_list_change_problem_number[i] + 1}번문제 문제저장용 파일이 존재하지 않습니다!")
        hwp.Open(rf"{problem_directory}")
        hwp.MoveToField(f"{src[1]}번문제")

        start_pos = hwp.GetPos()
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("MoveLeft")

        find_word(hwp, end_string, 10)
        hwp.HAction.Run("MoveLineEnd")
        end_pos1 = hwp.GetPos()

        hwp.SetPos(*start_pos)
        find_word(hwp, end_string2, 10)
        hwp.HAction.Run("MoveLineEnd")
        end_pos2 = hwp.GetPos()

        # 문제의 끝 정해주기
        if (end_pos1[0] == end_pos2[0]):
            if end_pos1[1] <= end_pos2[1]: end_pos = end_pos2
            else: end_pos = end_pos1
        elif end_pos1[0] < end_pos2[0]:
            end_pos = end_pos1
        else:
            end_pos = end_pos2
        
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.Run("DeleteBack")
        hwp.SetPos(*start_pos)
        hwp.Run("Paste")
        hwp.MoveToField(f"{src[1]}번문제")
        hwp.HAction.Run("StyleShortcut2")

        hwp.Save()
        myWindow.append_text_function_2(string=f"{field_list_change_problem_number[i] + 1}번문제 반영완료! ({i + 1}번째 입력)")
        progress += (100 // (problem_change_problem.shape[0] + problem_change_solution.shape[0]))
        myWindow.progressbar_2.setValue(progress)
        sleep(0.2)

    for i in range(problem_change_solution.shape[0]):
        problem_change_solution_set = problem_change_solution.iloc[i]
        src = array_to_problem_directory(problem_change_solution_set, grade=grade_number, test_name=test_name,
                                         for_release=False, ip_address=ip_address, test=test_mode)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        myWindow.append_text_function_2(string=f"{field_list_change_solution_number[i] + 1}번풀이 반영중...({i + 1}번째 입력)")
        hwp.Open(rf'{basefile}')
        if (hwp.EditMode == 0):
            solution_read_only.append(dst_problem_number)
            continue
        
        hwp.MoveToField(f"{field_list_change_solution_number[i] + 1}번풀이글상자", start=True)
        start_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.MoveToField(f"{field_list_change_solution_number[i] + 1}번풀이글상자", start=False)
        end_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.HAction.Run("Copy")
        hwp.HAction.Run("MoveDown")

        if not os.path.exists(problem_directory):
            raise Exception(f"{field_list_change_solution_number[i] + 1}번풀이 문제저장용 파일이 존재하지 않습니다!")
        hwp.Open(rf"{problem_directory}")
        hwp.MoveToField(f"{src[1]}번풀이")
        start_pos = hwp.GetPos()
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("MoveRight")
        end_pos = hwp.GetPos()
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.Run("DeleteBack")
        hwp.SetPos(*start_pos)
        hwp.Run("Paste")
        hwp.Save()
        myWindow.append_text_function_2(string=f"{field_list_change_solution_number[i] + 1}번풀이 반영완료! ({i + 1}번째 입력)")
        sleep(0.2)
        progress += (100 // (problem_change_problem.shape[0] + problem_change_solution.shape[0]))
        myWindow.progressbar_2.setValue(progress)

    end_time = dt.now()
    elapsed_time = end_time - start_time
    myWindow.append_text_function_2(f'반영을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')

    if (len(problem_read_only) != 0):
        myWindow.append_text_function_2(f'읽기 전용으로 열려 반영하지 못한 문제 번호 : ')
        myWindow.append_text_function_2(' '.join(map(lambda x: str(x)+'번', problem_read_only)))
        myWindow.append_text_function_2(f' ')
    
    if (len(solution_read_only) != 0):
        myWindow.append_text_function_2(f'읽기 전용으로 열려 반영하지 못한 풀이 번호 : ')
        myWindow.append_text_function_2(' '.join(map(lambda x: str(x)+'번', solution_read_only)))
        myWindow.append_text_function_2(f' ')
    myWindow.progressbar_2.setValue(100)
    hwp.Quit()

def add_reference_gui(hwp, test_name: str, grade_number, ip_address: str, reference_string: str, excel=None,
                      test_mode: bool = False):
    start_time = dt.now()
    if not test_mode:
        myWindow.append_text_function_3(string=f"{test_name} 출처 표시 시작")
    else:
        myWindow.append_text_function_3(string=f"{test_name} 출처 표시 시작 (테스트 모드)")

    progress = 0
    myWindow.progressbar_3.setValue(progress)
    myWindow.append_text_function_3(string="엑셀 로딩중...")
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    myWindow.append_text_function_3(string="엑셀 로딩 완료")

    dst_problem_number_for_field = [x for x in range(1, len(readexcel(excel, grade=grade_number)[test_name]) + 1)]
    problem_read_only = []
    for i in range(problems.shape[0]):
        myWindow.append_text_function_3(string=f"{dst_problem_number_for_field[i]}번 출처 표시중...({i + 1}번째 입력)")
        problem_set = problems.iloc[i]
        src = array_to_problem_directory(problem_set, grade=grade_number, test_name=test_name, for_release=False,
                                         test=test_mode, ip_address=ip_address)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]

        if not os.path.exists(problem_directory):
            raise Exception(f"{i + 1}번문제 문제저장용 파일이 존재하지 않습니다!")
        hwp.Open(rf"{problem_directory}")
        if (hwp.EditMode == 0):
            problem_read_only.append(dst_problem_number)
            continue

        hwp.MoveToField(f"{src[1]}번문제")
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("MoveRight")
        hwp.HAction.Run("BreakLine")

        start_pos = hwp.GetPos()
        start_line = hwp.KeyIndicator()[5]
        insert_text(hwp, string=reference_string)
        hwp.HAction.Run("StyleShortcut6")
        end_pos = hwp.GetPos()
        end_line = hwp.KeyIndicator()[5]
  
        while (start_line != end_line):
            hwp.SetPos(*start_pos)
            hwp.Run("Select")
            hwp.SetPos(*end_pos)
            hwp.HAction.Run("CharShapeSpacingDecrease")
            hwp.HAction.Run("CharShapeSpacingDecrease")
            hwp.HAction.Run("MoveRight")
            end_pos = hwp.GetPos()
            end_line = hwp.KeyIndicator()[5]

        hwp.Save()
        myWindow.append_text_function_3(string=f"{dst_problem_number_for_field[i]}번 출처 표시 완료! ({i + 1}번째 입력)")
        progress += (100 // problems.shape[0])
        myWindow.progressbar_3.setValue(progress)

    sleep(0.2)
    end_time = dt.now()
    elapsed_time = end_time - start_time
    myWindow.append_text_function(f'출처 표시를 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')
    if (len(problem_read_only) != 0):
        myWindow.append_text_function_2(f'읽기 전용으로 열려 출처를 표시하지 못한 문제 번호 : ')
        myWindow.append_text_function_2(' '.join(map(lambda x: str(x)+'번', problem_read_only)))
        myWindow.append_text_function_2(f' ')
    myWindow.progressbar.setValue(100)
    hwp.Quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    try:
        app.exec_()
    except:
        pass