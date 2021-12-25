# HWP파일에서 자주 사용하는 함수 정리
import os
import psutil

def find_word(hwp, word, size, direction="Forward"):
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir(direction)
    hwp.HParameterSet.HFindReplace.FindCharShape.Height = hwp.PointToHwpUnit(size)
    hwp.HParameterSet.HFindReplace.FindString = word
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HParameterSet.HFindReplace.SeveralWords = 1
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

def find_random_word(hwp, size, direction = "Forward"):
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindCharShape.Height = hwp.PointToHwpUnit(size)
    hwp.HParameterSet.HFindReplace.FindString = "\\k"
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir(direction)
    hwp.HParameterSet.HFindReplace.FindRegExp = 1
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindType = 1
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

def find_change_word(hwp, direction = "Forward"):
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindCharShape.TextColor = hwp.RGBColor(255, 0, 0)
    hwp.HParameterSet.HFindReplace.FindCharShape.Height = hwp.PointToHwpUnit(30.0)
    hwp.HParameterSet.HFindReplace.FindCharShape.FindString = "\\z"
    hwp.HParameterSet.HFindReplace.FindCharShape.Direction = hwp.FindDir(direction)
    hwp.HParameterSet.HFindReplace.FindCharShape.WholeWordOnly = 1
    hwp.HParameterSet.HFindReplace.FindCharShape.FindRegExp = 1
    hwp.HParameterSet.HFindReplace.FindCharShape.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.FindCharShape.FindType = 1
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

def add_field(hwp, field_name):
    hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
    hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = field_name
    hwp.HParameterSet.HInsertFieldTemplate.TemplateName = field_name
    hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)

def get_current_keyindicator(hwp, source : str):
    hwp.Open(f"{source}")
    keyindicator = hwp.KeyIndicator()
    return keyindicator

def get_current_pos(hwp, source : str):
    hwp.Open(f"{source}")
    current_position = hwp.GetPos()
    return current_position

def circle_word(hwp, string : str):
    hwp.HAction.GetDefault("ComposeChars", hwp.HParameterSet.HChCompose.HSet)
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeHangul = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeLatin = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeHanja = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeJapanese = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeOther = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeSymbol = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CharShapes.CircleCharShape.FontTypeUser = hwp.FontType("TTF")
    hwp.HParameterSet.HChCompose.CircleType = 1
    hwp.HParameterSet.HChCompose.CheckCompose = 0
    hwp.HParameterSet.HChCompose.Chars = string
    hwp.HAction.Execute("ComposeChars", hwp.HParameterSet.HChCompose.HSet)

def insert_text(hwp, string : str):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = string
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def shape_copy_paste(hwp, type = 2):
    hwp.HAction.GetDefault("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
    hwp.HParameterSet.HShapeCopyPaste.type = type  # 글자 모양과 문단 모양 둘 다 복사
    hwp.HAction.Execute("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("StyleShortcut1")
    hwp.HAction.GetDefault("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
    hwp.HAction.Execute("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)

def insert_file(hwp, FileName, KeepSection = 1, KeepCharshape = 1, KeepParashape = 1, KeepStyle = 1):
    """
    :param hwp: hwp 파일 기본
    :param FileName: 파일 경로
    :param KeepSection: 글자 모양 유지
    :param KeepCharshape: 쪽 모양 유지
    :param KeepParashape: 문단 모양 유지
    :param KeepStyle: 스타일 유지
    0의 경우 유지 안함, 1의 경우 유지
    """
    hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
    hwp.HParameterSet.HInsertFile.filename = FileName
    hwp.HParameterSet.HInsertFile.KeepSection = KeepSection
    hwp.HParameterSet.HInsertFile.KeepCharshape = KeepCharshape
    hwp.HParameterSet.HInsertFile.KeepParashape = KeepParashape
    hwp.HParameterSet.HInsertFile.KeepStyle = KeepStyle
    hwp.HAction.Execute("InsertFile", hwp.HParameterSet.HInsertFile.HSet)

def move_to_field(hwp, field, text = True, start = True, select = False):
    """
    :param hwp: hwp 파일
    :param field: 필드 이름, 중복된 필드 이름이 있을 경우 이름 뒤에 '{{#}}'로 번호를 지정할 수 있다.
    :param text: True시 누름틀 내부, False시 누름틀 코드로 이동, 기본값은 True
    :param start: True시 필드의 처음, False시 필드의 끝으로 이동, 기본값은 True
    :param select: True시 필드 선택시 블록 단위, False시 캐럿만 이동
    :return: None
    """
    hwp.MoveToField(f"{field}", text = text, start = start, select = select)

def get_field_list(hwp):
    field_list = hwp.GetFieldList().split("\x02")
    return field_list

def put_field_text(hwp, field: str, text: str):
    hwp.PutFieldText(Field = field, Text=text)

def kill_hwp():
    my_pid = os.getpid()
    for proc in psutil.process_iter():
        try:
            process_name = proc.name()
            if proc.pid == my_pid:
                continue

            if (process_name.startswith('Hwp')) and ('Automation' in (''.join(proc.cmdline()))):
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            print(e)

if __name__ == "__main__":
    import win32com.client as win32
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    get_current_keyindicator()