import win32com.client as win32
from source_to_basefile import *
def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.HAction.GetDefault("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    hwp.HParameterSet.HViewProperties.ZoomType = hwp.HwpZoomType("FitPage")
    hwp.HAction.Execute("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    return hwp
hwp = init_hwp()
"""
임시로 이 파일을 이용해서 작업해주세요.
문의사항은
- a01032208149@gmail.com
또는
- 010-3220-8149로 연락주세요.
"""

# 1. 주문서(엑셀)을 통해 검토용파일 제작하기
source_to_problem_execute(hwp = hwp, excel = r"C:\Users\Season\Desktop\준호타이핑용\testbench\태풍\시험1.xls", grade_number = 2, test_name = "1번 함수의 극한~함수의 그래프 (210828)")

#2. 검토용파일 수정 이후 반영 및 출처표시하기
# basefile_to_source(hwp = hwp, basefile = r"C:\Users\Season\Desktop\준호타이핑용\testbench\태풍\testbench_basefiletest.hwp", grade_number = 1, excel = r"C:\Users\Season\Desktop\내신주문서.xlsx")