import sys
import win32com.client as win32
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtGui import QIcon
from source_to_basefile import *

def new_basefile_gui(file_name : str):
    source_directory = r"D:\PythonTyphoon\태풍\기출_문제+답지_원본_2문제씩_번호o.hwp"
    shutil.copyfile(source_directory, rf'D:\PythonTyphoon\태풍\{file_name}_검토용파일_(문제+답지).hwp')
    new_file = rf'D:\PythonTyphoon\태풍\{file_name}_검토용파일_(문제+답지).hwp'
    return new_file

def new_basefile_no_number_gui(file_name : str):
    source_directory = r"D:\PythonTyphoon\태풍\기출_문제+답지_원본_2문제씩_번호x.hwp"
    shutil.copyfile(source_directory, rf'D:\PythonTyphoon\태풍\{file_name}_검토용파일_(문제+답지).hwp')
    new_file = rf'D:\PythonTyphoon\태풍\{file_name}_검토용파일_(문제+답지).hwp'
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
        self.ui = uic.loadUi("test.ui", self)
        self.setWindowTitle("검토용파일 제작 프로그램")
        self.setWindowIcon(QIcon("icon.png"))
        self.pushButton_execute.clicked.connect(self.execute_function)

    def grade_function(self):
        if self.radioButton_grade1.isChecked() :  grade_number = 1
        elif self.radioButton_grade2.isChecked() : grade_number = 2

    def getText_excel_directory(self):
        excel_directory = self.QTextEdit_excel_directory.toPlainText()

    def getText_test_name(self):
        test_name = self.QTextEdit_test_name.toPlainText()

    def appendTextFunction(self, string) :
        QCoreApplication.processEvents()
        self.textBrowser_progress.append(rf"{string}")

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

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    try:
        app.exec_()
    except:
        myWindow.appendTextFunction("종료중...")