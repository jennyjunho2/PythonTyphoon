import sys
import win32com.client as win32
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon
from source_to_basefile import *

def new_basefile_gui(file_name : str, grade_number):
    file_name_date = file_name[file_name.find("(")+1:file_name.find(")")]
    file_name_count = file_name[file_name.find("번")-2] if '번' in file_name else 0
    file_name = file_name.replace(f"({file_name_date})", "").replace("")
    if file_name_count != 0:
        file_copy_directory = str(file_name_date)+"_" + str(grade_number) +" "+ str(file_name_count) + file_name + str("_검토용파일.hwp")
    else:
        file_copy_directory = str(file_name_date)+"_" +str(grade_number) + " " + file_name + str("_검토용파일.hwp")
    source_directory = r"D:\PythonTyphoon\태풍\기출_문제+답지_원본_2문제씩_번호o.hwp"
    shutil.copyfile(source_directory, file_copy_directory)
    new_file = file_copy_directory
    return new_file

def new_basefile_no_number_gui(file_name : str, grade_number):
    file_name_date = file_name[file_name.find("(") + 1:file_name.find(")")]
    file_name_count = file_name[file_name.find("번") - 2] if '번' in file_name else 0
    file_name = file_name.replace(f"({file_name_date})", "").replace("")
    if file_name_count != 0:
        file_copy_directory = str(file_name_date) + "_" + str(grade_number) + " " + str(file_name_count) + file_name + str("_검토용파일.hwp")
    else:
        file_copy_directory = str(file_name_date) + "_" + str(grade_number) + " " + file_name + str("_검토용파일.hwp")
    source_directory = r"D:\PythonTyphoon\태풍\기출_문제+답지_원본_2문제씩_번호x.hwp"
    shutil.copyfile(source_directory, file_copy_directory)
    new_file = file_copy_directory
    return new_file

def source_to_problem_execute_gui(hwp, excel : str, grade_number : int, test_name : str, basefile :bool = True):
    start_time = dt.now()
    myWindow.appendTextFunction(string = f"{test_name} 제작 시작")
    progress = 0
    myWindow.progressbar.setValue(progress)
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    dst_problem_number_for_field = [x for x in range(1, len(readexcel(excel, grade = grade_number)[test_name])+1)]
    dst = new_basefile_gui(test_name) if basefile == False else new_basefile_no_number_gui(test_name)
    for i in range(problems.shape[0]):
        myWindow.appendTextFunction(string=f"{dst_problem_number_for_field[i]}번 입력중...({i + 1}번째 입력)")
        problem_set = problems.iloc[i]
        src = array_to_problem_directory(problem_set, grade=grade_number, test_name = test_name)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        source_to_basefile_problem(hwp, source = problem_directory , source_number = src_problem_number, destination = dst, destination_number = dst_problem_number_for_field[i])
        source_to_basefile_solution(hwp, source = problem_directory, source_number = src_problem_number, destination = dst, destination_number = dst_problem_number_for_field[i])
        hwp.PutFieldText(Field = f"{i+1}번문제번호", Text = str(replace_number_to_question[int(dst_problem_number)]) if int(dst_problem_number) >= 41 else str(int(dst_problem_number)))
        hwp.PutFieldText(Field = f"{i+1}번풀이번호", Text = str(replace_number_to_question[int(dst_problem_number)]) if int(dst_problem_number) >= 41 else str(int(dst_problem_number)))
        myWindow.appendTextFunction(string = f"{dst_problem_number_for_field[i]}번 입력완료! ({i+1}번째 입력완료)")
        hwp.Save()
        progress += (100 // problems.shape[0])
        myWindow.progressbar.setValue(progress)
    if basefile == True:
        hwp.PutFieldText(Field = "검토용파일이름", Text = test_name)

    hwp.Save()
    sleep(0.2)
    end_time = dt.now()
    elapsed_time = end_time - start_time
    myWindow.appendTextFunction(f'입력을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')
    myWindow.progressbar.setValue(100)
    hwp.Quit()

def basefile_to_source_gui(hwp, basefile : str, grade_number, excel = None, reference : bool = False):
    start_time = dt.now()
    hwp.Open(rf'{basefile}')
    # 검토용파일 존재하는지 검사
    if os.path.exists(basefile) == False:
        raise Exception("검토용파일이 존재하지 않습니다!")
    myWindow.appendTextFunction_2(string=f"{basefile} 반영 시작")
    progress = 0
    myWindow.progressbar_2.setValue(progress)
    test_name = hwp.GetFieldText("검토용파일이름")
    myWindow.appendTextFunction_2(string = test_name+" 반영 진행중...")
    problems = get_problem_list(excel=excel, grade=grade_number, test_name=test_name)
    field_list = hwp.GetFieldList().split("\x02")
    field_list_problem_number = [x for x in field_list if "번문제번호" in x]
    field_list_solution_number = [x for x in field_list if "번풀이번호" in x]
    field_list_change_problem_number = []
    field_list_change_solution_number = []

    for field_problem_number in field_list_problem_number:
        hwp.MoveToField(field_problem_number, start = False)
        hwp.HAction.Run("SelectAll")
        if hwp.CharShape.Item("TextColor") == 255: # 빨간색일 경우
            hwp.HAction.Run("MoveLeft")
            field_list_change_problem_number.append(hwp.GetCurFieldName())
        else:
            pass

    for field_solution_number in field_list_solution_number:
        hwp.MoveToField(field_solution_number, start=False)
        hwp.HAction.Run("SelectAll")
        if hwp.CharShape.Item("TextColor") == 255: # 빨간색일 경우
            hwp.HAction.Run("MoveLeft")
            field_list_change_solution_number.append(hwp.GetCurFieldName())
        else:
            pass
    field_list_change_problem_number = list(map(lambda y : int(y)-1, list(map(lambda x: x[:-5], field_list_change_problem_number))))
    field_list_change_solution_number = list(map(lambda y : int(y)-1, list(map(lambda x: x[:-5], field_list_change_solution_number))))
    problem_change_problem = problems.iloc[field_list_change_problem_number]
    problem_change_solution = problems.iloc[field_list_change_solution_number]
    for i in range(problem_change_problem.shape[0]):
        problem_change_problem_set = problem_change_problem.iloc[i]
        src = array_to_problem_directory(problem_change_problem_set, grade=grade_number, test_name = test_name)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        myWindow.appendTextFunction_2(string = f"{field_list_change_problem_number[i]+1}번문제 반영중...({i+1}번째 입력)")
        hwp.Open(rf'{basefile}')
        hwp.MoveToField(f"{field_list_change_problem_number[i]+1}번문제", start = True)
        start_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.MoveToField(f"{field_list_change_problem_number[i]+1}번문제", start = False)
        end_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.HAction.Run("Copy")
        hwp.HAction.Run("MoveDown")

        hwp.Open(rf"{problem_directory}")
        if os.path.exists(problem_directory) == False:
            raise Exception(f"{field_list_change_problem_number[i]+1}번문제 문제저장용 파일이 존재하지 않습니다!")
        hwp.MoveToField(f"{src[1]}번문제")
        hwp.HAction.Run("MoveRight")
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
        myWindow.appendTextFunction_2(string = f"{field_list_change_problem_number[i] + 1}번문제 반영완료! ({i + 1}번째 입력)")
        progress += (100 // (problem_change_problem.shape[0]+problem_change_solution))
        myWindow.progressbar_2.setValue(progress)
        sleep(0.2)

    for i in range(problem_change_solution.shape[0]):
        problem_change_solution_set = problem_change_solution.iloc[i]
        src = array_to_problem_directory(problem_change_solution_set, grade=grade_number, test_name = test_name)
        problem_directory, src_problem_number, dst_problem_number, src_problem_score = src[0], src[1], src[2], src[3]
        myWindow.appendTextFunction_2(string = f"{field_list_change_solution_number[i]+1}번풀이 반영중...({i+1}번째 입력)")
        hwp.Open(rf'{basefile}')
        hwp.MoveToField(f"{field_list_change_solution_number[i]+1}번풀이", start = True)
        start_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.MoveToField(f"{field_list_change_solution_number[i]+1}번풀이", start = False)
        end_pos = hwp.GetPos()
        hwp.Run("Cancel")
        hwp.SetPos(*start_pos)
        hwp.Run("Select")
        hwp.SetPos(*end_pos)
        hwp.HAction.Run("Copy")
        hwp.HAction.Run("MoveDown")

        hwp.Open(rf"{problem_directory}")
        if os.path.exists(problem_directory) == False:
            raise Exception(f"{field_list_change_solution_number[i]+1}번풀이 문제저장용 파일이 존재하지 않습니다!")
        hwp.MoveToField(f"{src[1]}번풀이")
        hwp.HAction.Run("MoveRight")
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
        myWindow.appendTextFunction_2(string = f"{field_list_change_solution_number[i] + 1}번풀이 반영완료! ({i + 1}번째 입력)")
        sleep(0.2)
        progress += (100 // (problem_change_problem.shape[0] + problem_change_solution))
        myWindow.progressbar_2.setValue(progress)

    end_time = dt.now()
    elapsed_time = end_time - start_time
    myWindow.appendTextFunction_2(f'반영을 완료하였습니다. 약 {elapsed_time.seconds}초 소요되었습니다.')
    myWindow.progressbar_2.setValue(100)
    hwp.Quit()

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = False
    hwp.HAction.GetDefault("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    hwp.HParameterSet.HViewProperties.ZoomType = hwp.HwpZoomType("FitPage")
    hwp.HAction.Execute("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    return hwp
hwp = init_hwp()

class WindowClass(QDialog) :
    def __init__(self) :
        super().__init__()
        # self.ui = uic.loadUi(r"C:\Users\Season\Desktop\준호타이핑용\testbench\typhoon_gui\test.ui", self)
        self.ui = uic.loadUi("test.ui", self)
        self.setWindowTitle("검토용파일 제작 프로그램")
        self.setWindowIcon(QIcon("icon.png"))
        self.pushButton_execute.clicked.connect(self.execute_function)
        self.pushButton_execute_2.clicked.connect(self.execute_function_2)
        self.pushButton_find_excel.clicked.connect(self.get_save_file_name)
        self.pushButton_find_excel_2.clicked.connect(self.get_save_file_name_2)


    def getText_excel_directory(self):
        excel_directory = self.QTextEdit_excel_directory.toPlainText()

    def getText_test_name(self):
        test_name = self.QTextEdit_test_name.toPlainText()

    def appendTextFunction(self, string) :
        QCoreApplication.processEvents()
        self.textBrowser_progress.append(rf"{string}")

    def get_save_file_name(self):
        file_filter = 'Data File (*.xlsx *.xls)'
        response = QFileDialog.getSaveFileName(
            parent = self,
            caption = "엑셀 파일을 선택해주세요.",
            filter = file_filter,
            initialFilter = "Excel File (*.xlsx *.xls)"
        )
        self.QTextEdit_excel_directory.setPlainText(response[0])

    def execute_function(self):
        if self.radioButton_grade1.isChecked() :  grade_number = 1
        elif self.radioButton_grade2.isChecked() : grade_number = 2
        excel_directory = self.QTextEdit_excel_directory.toPlainText().strip('""')
        test_name = self.QTextEdit_test_name.toPlainText().strip('""')
        try:
            self.pushButton_execute.setEnabled(False)
            source_to_problem_execute_gui(hwp = hwp,excel = excel_directory, test_name = test_name, grade_number= grade_number)
            self.pushButton_execute.setEnabled(True)
            hwp.Quit()
        except OSError:
            myWindow.appendTextFunction(string = "주문서 경로를 확인해주세요.")
            self.pushButton_execute.setEnabled(True)
            hwp.Quit()
        except UnboundLocalError:
            myWindow.appendTextFunction(string = "학년을 선택해주세요.")
            self.pushButton_execute.setEnabled(True)
            hwp.Quit()
        except ValueError:
           myWindow.appendTextFunction(string = "학년과 시험지 이름을 확인해주세요.")
           self.pushButton_execute.setEnabled(True)
           hwp.Quit()
        except Exception as e:
           myWindow.appendTextFunction(string="오류가 발생했습니다 : " + str(e))
           self.pushButton_execute.setEnabled(True)
           hwp.Quit()

################################################################################################################

    def getText_excel_directory_2(self):
        basefile_directory = self.QTextEdit_excel_directory_2.toPlainText()

    def getText_test_name_2(self):
        test_name = self.QTextEdit_test_name_2.toPlainText()

    def appendTextFunction_2(self, string):
        QCoreApplication.processEvents()
        self.textBrowser_progress_2.append(rf"{string}")

    def get_save_file_name_2(self):
        file_filter = '한글 파일 (*.hwp)'
        response = QFileDialog.getSaveFileName(
            parent = self,
            caption = "한글 파일을 선택해주세요.",
            filter = file_filter,
            initialFilter = "hwp File (*.hwp)"
        )
        self.QTextEdit_excel_directory_2.setPlainText(response[0])

    def execute_function_2(self):
        if self.radioButton_grade1_2.isChecked() :  grade_number = 1
        elif self.radioButton_grade2_2.isChecked() : grade_number = 2
        excel_directory = self.QTextEdit_excel_directory_2.toPlainText().strip('""')
        basefile_directory = self.QTextEdit_test_name_2.toPlainText().strip('""')
        try:
            self.pushButton_execute.setEnabled(False)
            basefile_to_source_gui(hwp = hwp, excel = excel_directory, basefile = basefile_directory, grade_number= grade_number)
            self.pushButton_execute_2.setEnabled(True)
            hwp.Quit()
        except OSError:
            myWindow.appendTextFunction_2(string = "검토용파일 경로를 확인해주세요.")
            self.pushButton_execute_2.setEnabled(True)
            hwp.Quit()
        except UnboundLocalError:
            myWindow.appendTextFunction_2(string = "학년을 선택해주세요.")
            self.pushButton_execute_2.setEnabled(True)
            hwp.Quit()
        except ValueError:
           myWindow.appendTextFunction_2(string = "학년과 시험지 이름을 확인해주세요.")
           self.pushButton_execute_2.setEnabled(True)
           hwp.Quit()
        except Exception as e:
           myWindow.appendTextFunction_2(string="오류가 발생했습니다 : " + str(e))
           self.pushButton_execute_2.setEnabled(True)
           hwp.Quit()



if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    try:
        app.exec_()
    except:
        myWindow.appendTextFunction("종료중...")