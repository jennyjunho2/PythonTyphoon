import win32com.client as win32
from time import sleep
from source_to_basefile import *
from basefile_to_exam import *

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.HAction.GetDefault("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    hwp.HParameterSet.HViewProperties.ZoomType = hwp.HwpZoomType("FitPage")
    hwp.HAction.Execute("ViewZoom", hwp.HParameterSet.HViewProperties.HSet)
    return hwp

hwp = init_hwp()
src = r"D:\PythonTyphoon\태풍\testbench_src.hwp"
dst = r"D:\PythonTyphoon\태풍\testbench_dst.hwp"
for i in range(1, 5):
    source_to_basefile_problem(hwp, src, i, dst, i)
    source_to_basefile_solution(hwp, src, i, dst, i)
add_score(hwp, dst, 1, 3)
add_score(hwp, dst, 2, 4)
add_score(hwp, dst, 3, 5)
add_score(hwp, dst, 4, 6)

