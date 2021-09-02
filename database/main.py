import win32com.client as win32
from time import sleep

def init_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp
def extract_eqn(hwp):  # 이전 포스팅에서 소개한, 수식 추출방법을 함수로 정의
    Act = hwp.CreateAction("EquationModify")
    Set = Act.CreateSet()
    Pset = Set.CreateItemSet("EqEdit", "EqEdit")
    Act.GetDefault(Pset)
    return Pset.Item("String")

hwp = init_hwp()
hwp.Open(r"D:\PythonTyphoon\database\testbench_databasetest.hwp")
eqn_dict = {}  # 사전 형식의 자료 생성 예정
ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.

while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
    nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
    if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
        position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
        position = position.Item("List"), position.Item("Para"), position.Item("Pos")
        hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
        hwp.FindCtrl()  # 해당 컨트롤 선택
        eqn_string = extract_eqn(hwp)  # 문자열 추출
        eqn_dict[position] = eqn_string  # 좌표가 key이고, 수식문자열이 value인 사전 생성
    ctrl = nextctrl  # 다음 컨트롤 탐색
hwp.Run("Cancel")  # 완료했으면 선택해제

"""dict를 가지고 수식 삽입하기"""

for position, eqn_string in eqn_dict.items():
    hwp.HAction.GetDefault("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)
    hwp.HParameterSet.HEqEdit.EqFontName = "HancomEQN"
    hwp.HParameterSet.HEqEdit.string = eqn_string
    hwp.HParameterSet.HEqEdit.BaseUnit = hwp.PointToHwpUnit(10.0)  # 수식 폰트 크기 : 30
    hwp.HAction.Execute("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)  # 폰트이상함
    sleep(1)  # 시연을 위해 1초 멈춤
    hwp.FindCtrl()  # 다시 선택
    hwp.HAction.GetDefault("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 3)
    hwp.HParameterSet.HShapeObject.Version = "Equation Version 60"
    hwp.HParameterSet.HShapeObject.EqFontName = "HancomEQN"
    hwp.HParameterSet.HShapeObject.HSet.SetItem("ApplyTo", 0)
    hwp.HParameterSet.HShapeObject.HSet.SetItem("TreatAsChar", 1)
    hwp.HAction.Execute("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.Run("Cancel")  # 폰트 예뻐짐
    hwp.Run("MoveRight")  # 다음 수식 삽입 준비
    sleep(1)  # 시연을 위해 1초 멈춤

for key, value in eqn_dict.items():
    print(f"{key} : {value}")