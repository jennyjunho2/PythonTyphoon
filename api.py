# HWP파일에서 자주 사용하는 함수 정리
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

if __name__ == "__main__":
    import win32com.client as win32
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    get_current_keyindicator()