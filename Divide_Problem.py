import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
import os, shutil

import win32com.client as win
import fnmatch
import time
from PIL import Image
import olefile
import pyautogui
import re




hwp=win.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") #한글 권한 설정 보안 레지스트리
hwp.XHwpWindows.Item(0).Visible=False #백그라운드에서 동작(False) / 보이면서 동작(True)
hwp.SetMessageBoxMode(0x00020000) # 예/아니오 -> 아니오
hwp.SetMessageBoxMode(0x00000001) # 확인 박스 -> OK
# hwp.SetMessageBoxMode(0x00000010) # 확인/취소 박스 -> OK


root = Tk()
root.title("Postmath_Divide_HWP")

# 한글 문제 해설 정답 쪼개기 관련 함수 모음

def find_mizu(): # 미주 찾아가기
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 31)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)


def find_mizunum(): # 미주 번호 찾아가기
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 32)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)


def count_mizu(): # 미주 개수 세기 -> 변수 cnt_mizu 에 저장
    hwp.MovePos(2) # 문서 제일 앞으로
    find_mizunum() # 미주번호 찾아가 : hwp.MovePos(2) 이거 때매 맨 첫 미주번호로
    global cnt_mizu # 미주개수를 전역 변수로
    cnt_mizu = 1
    while hwp.HAction.Run("NoteToNext") == True:
        cnt_mizu = cnt_mizu + 1
    hwp.MovePos(2) # 문서 제일 앞으로
    


def onepageoneproblem(): # 한쪽에 한문제 (첫페이지는 표지)
    hwp.MovePos(2)
    i=0
    while i < cnt_mizu :
        find_mizu()
        hwp.HAction.Run("BreakPage"); # ctrl+enter
        hwp.HAction.Run("MoveRight");
        i = i+1
    hwp.HAction.Run("MoveColumnEnd"); #단의 끝점으로 이동
    hwp.HAction.Run("BreakPage");
    hwp.MovePos(3) # 문서 제일 뒤로
    hwp.HAction.Run("MoveLeft");
    hwp.HAction.Run("BreakPage");
    hwp.HAction.Run("MoveRight");
    hwp.HAction.Run("BreakPage");
    hwp.MovePos(2) # 문서 제일 앞으로


def page_size_set(): # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)
    hwp.HParameterSet.HSecDef.PageDef.PaperWidth = hwp.MiliToHwpUnit(106.5)
    hwp.HParameterSet.HSecDef.PageDef.PaperHeight = hwp.MiliToHwpUnit(1110.0)
    hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(0.0)
    hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)


def multicolumn_1(): # 페이지 1단으로 
    hwp.HAction.GetDefault("MultiColumn", hwp.HParameterSet.HColDef.HSet)
    hwp.HParameterSet.HColDef.Count = 1
    hwp.HParameterSet.HColDef.SameGap = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyClass", 832)
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyTo", 6)
    hwp.HAction.Execute("MultiColumn", hwp.HParameterSet.HColDef.HSet)



def tabdiv_pro_sol(): # 1탭에 문제만 2탭에 해설만 / 작동 후 해설 맨 처음
    # 새 탭 만들고 다시 돌아와
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("WindowNextTab")


    # 미주 찾아가서 그 미주있는 페이지 복사 후 새탭에 넣어 (새탭1)
    # hwp.MovePos(2) # 문서 제일 앞으로
    find_mizu()
    hwp.HAction.Run("MoveSelPageDown")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    time.sleep(0.1)
    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("PasteOriginal")
    time.sleep(0.1)
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)


    # 미주 번호 뒤에 있는 내용(해설내용) 복사 후 새탭에 넣어 (새탭2)
    hwp.MovePos(2)
    find_mizunum()
    hwp.HAction.Run("SelectAll")
    time.sleep(0.1)
    hwp.HAction.Run("Copy")
    time.sleep(0.1)
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("Paste")
    time.sleep(0.1)
    hwp.MovePos(2)
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("DeleteBack")
    
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)


    # 문제 부분 미주 없애고 해설 페이지로 돌아가
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("WindowNextTab")
    hwp.MovePos(2)
    find_mizu()
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("DeleteBack")
    hwp.HAction.Run("MoveViewEnd");
    hwp.HAction.Run("MoveSelTopLevelEnd");
    hwp.HAction.Run("Delete");
    hwp.HAction.Run("WindowNextTab")
    time.sleep(0.1)
    hwp.MovePos(3) # 해설 맨 마지막 끝으로가서
    
    global solution_page
    solution_page = hwp.KeyIndicator()[3] # 현재 커서의 페이지 번호 저장
    
    hwp.MovePos(2) # 해설 맨 처음으로


def image_merge(): # 해설 이미지 여러쪽이면 병합하고 남는거 지우기
    im_list = []
    for i in range(1, solution_page+1):
        im_list.append(sol_png_name+str(i).zfill(3)+".png")


    images = [Image.open(x) for x in im_list]

    # size -> size[0] : width, size[1] : height
    widths = [x.size[0] for x in images]
    heights = [x.size[1] for x in images]

    # 가장 width가 넓고 heights를 다 합친 큰 스케치북을 준비
    # 먼저 최대 넓이, 전체 높이 구해오기
    max_width, total_height = max(widths), sum(heights)

    # 스케치북 준비
    result_img = Image.new("RGB", (max_width, total_height)) 
    y_offset = 0 # y 위치 정보
    for img in images:
        result_img.paste(img, (0, y_offset))
        y_offset += img.size[1] # height 값 만큼 더해줌
    result_img.save(sol_png_name+".png", "png")
    time.sleep(0.2)
    for i in range(1, solution_page+1):
        os.remove(sol_png_name+str(i).zfill(3)+".png")


def save_sol_hml(num):  # 해설 hml파일 저장할거야 num = 문제번호
    global sol_hml_name
    sol_hml_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [1hml]" # 확장자 없는 이름
    while os.path.isfile(sol_hml_name+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = sol_hml_name+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_sol_png(num): # 해설 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global sol_png_name
    sol_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [2해설] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = sol_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)


def save_sol_hwp(num):  # 해설 hwp파일 저장할거야 num = 문제번호
    global sol_hwp_name
    sol_hwp_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [3hwp]" # 확장자 없는 이름
    while os.path.isfile(sol_hwp_name+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = sol_hwp_name+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)


def save_pro_hml(num):  # 문제 hml파일 저장할거야 num = 문제번호
    global pro_hml_name
    pro_hml_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [1문제] [1hml]"
    while os.path.isfile(pro_hml_name+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = pro_hml_name+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_pro_png(num): # 문제 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global pro_png_name
    pro_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [1문제] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = pro_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)
    
    os.rename(pro_png_name+"001"+".png", pro_png_name+".png")

def save_presol_hwp(num):  # 정답 hwp파일 저장할거야 num = 문제번호
    global presol_hwp_name
    presol_hwp_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [3정답] [3hwp]" # 확장자 없는 이름
    while os.path.isfile(presol_hwp_name+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = presol_hwp_name+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_presol_png(num): # 정답 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global presol_png_name
    presol_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [3정답] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = presol_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)
    
    os.rename(presol_png_name+"001"+".png", presol_png_name+".png")



def find_equation(): #내 옆 수식을 텍스트로
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 37)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
    



def extract_eqn(hwp):  # 수식을 선택한 후 수식안 텍스트 추출
    Act = hwp.CreateAction("EquationModify")
    Set = Act.CreateSet()
    Pset = Set.CreateItemSet("EqEdit", "EqEdit")
    Act.GetDefault(Pset)
    return Pset.Item("String") # Pset.Item("String") 여기에 수식 안 내용 저장됨


"""모든 수식 텍스트 차례로 dict로 얻기.
키는 (List, Para, Pos), 값은 eqn_string"""

def equation_to_text_all(hwp): # 한글 파일 열어서 모든 수식 text화
    eqn_dict = {}  # 사전 형식의 자료 생성 예정 ##
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    global sum_reST
    sum_reST = [] # reST들을 하나의 리스트에 담기

    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = position.Item("List"), position.Item("Para"), position.Item("Pos")
            hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            Act = hwp.CreateAction("EquationModify")
            Set = Act.CreateSet()
            Pset = Set.CreateItemSet("EqEdit", "EqEdit")
            Act.GetDefault(Pset) # Pset.Item("String") 여기에 수식 안 내용 저장됨
            
            #임시 해보기
            ST = re.sub('`|~|\s|\n|rm|it|\{|\}', "", Pset.Item("String"))
            ST1 = re.sub('`|~|\s|\n', "", Pset.Item("String"))
            ST2 = re.sub('\{*rm\{*(\d+)\}*\}*', r'\1', ST1)
            
            reST = re.findall('[^\d+-]', ST)
            for i in reST:
                sum_reST.append(i)

            hwp.HAction.Run("Delete") ## 여기부터는 한글에서 수식을 다 텍스트로

            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            # hwp.HParameterSet.HInsertText.Text = Pset.Item("String");
            hwp.HParameterSet.HInsertText.Text = ST2;
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            eqn_dict[position] = Pset.Item("String")  # 좌표가 key이고, 수식문자열이 value인 사전 생성 ##
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.Run("Cancel")  # 완료했으면 선택해제
    # print(sum_reST)

    # for key, value in eqn_dict.items(): ##
    #     print(f"{key} : {value}")

def equation_to_text_first(hwp): # 한글 파일 열어서 첫수식 text화
    eqn_dict = {}  # 사전 형식의 자료 생성 예정 ##
    
    hwp.MovePos(2)
    
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.

    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = position.Item("List"), position.Item("Para"), position.Item("Pos")
            hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            Act = hwp.CreateAction("EquationModify")
            Set = Act.CreateSet()
            Pset = Set.CreateItemSet("EqEdit", "EqEdit")
            Act.GetDefault(Pset) 
            hwp.HAction.Run("Delete") ## 여기부터는 한글에서 수식을 다 텍스트로
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = Pset.Item("String");
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            eqn_dict[position] = Pset.Item("String")  # 좌표가 key이고, 수식문자열이 value인 사전 생성 ##
            break
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.Run("Cancel")  # 완료했으면 선택해제

    # for key, value in eqn_dict.items(): ##
    #     print(f"{key} : {value}")


def count_eqed(hwp): # 수식개수 세기 / cnt_eqed 라는 변수에 저장
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    global cnt_eqed
    cnt_eqed = 0
    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            cnt_eqed = cnt_eqed+1
        ctrl = nextctrl  # 다음 컨트롤 탐색
    # print("수식개수는 " + str(cnt_eqed))



def preview_sol_hwp(): # 빠른정답만들기(각파일로) (새탭이 없을때 사용해야함.)
    hwp.MovePos(2)
    for i in range(1, cnt_mizu+1):
    # for i in range(1, 2):
        find_mizunum()
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Copy");
        hwp.HAction.Run("Cancel");
        hwp.HAction.Run("FileNewTab")
        hwp.HAction.Run("Paste")
        hwp.HAction.Run("PasteOriginal")
        hwp.MovePos(2)
        find_mizunum()
        hwp.HAction.Run("MoveSelRight")
        hwp.HAction.Run("DeleteBack")
        save_presol_png(i) # 정답 이미지(png)파일 저장할거야
        count_eqed(hwp) # 수식개수 세기 / cnt_eqed 라는 변수에 저장
        
        if cnt_eqed == 0: # 만약 수식 개수가 0이면
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); # 엔터없애기
            hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
            hwp.HParameterSet.HFindReplace.FindString = "^n";
            hwp.HParameterSet.HFindReplace.ReplaceString = "";
            hwp.HParameterSet.HFindReplace.FindRegExp = 1;
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
            
            hwp.HAction.Run("SelectAll")
            text = hwp.GetTextFile("TEXT","saveblock"); # text란 변수에 전체 텍스트 넣기
            hwp.HAction.Run("Cancel");
            cnt_circ = len(re.findall(r'[①②③④⑤⑥⑦⑧⑨]', text)) # text 중 원문자 개수
            cnt_OX = len(re.findall(r'[OX]', text)) # text 중 대문자 O, X 의 개수
            # print(cnt_circ)
            hwp.MovePos(2)
            hwp.HAction.Run("BreakPara")
            hwp.MovePos(2)
            
            if cnt_circ == 1: # 원문자가 1개이면 
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                hwp.HParameterSet.HInsertText.Text = "[객관식(단일)]";
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            elif cnt_circ > 1: # 원문자가 2개 이상이면 
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                hwp.HParameterSet.HInsertText.Text = "[객관식(다중)]";
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            else: # 원문자가 없으면 
                if cnt_OX >= 1:  # OX 가 존재하면 
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                    hwp.HParameterSet.HInsertText.Text = "[객관식(OX)]";
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
                else:  # OX 가 없으면
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                    hwp.HParameterSet.HInsertText.Text = "[주관식(보기)]";
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
                    

        elif cnt_eqed == 1 : # 만약 수식 개수가 1이면 -> 일단 주관식
            equation_to_text_all(hwp)
            hwp.MovePos(2)
            hwp.HAction.Run("BreakPara")
            hwp.MovePos(2)
            if len(sum_reST) == 0: # +-숫자 외의 문자가 없으면 -> 정수
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                hwp.HParameterSet.HInsertText.Text = "[주관식(정수)]";
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            else : # +-숫자 외의 문자가 있으면 -> 정수 X
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                hwp.HParameterSet.HInsertText.Text = "[주관식(정수) X]";
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
        
        else : # 만약 수식 개수가 2 이상이면 -> [주관식(보기)]
            equation_to_text_all(hwp)
            hwp.MovePos(2)
            hwp.HAction.Run("BreakPara")
            hwp.MovePos(2)
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "[주관식(보기)]";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
        
        save_presol_hwp(i) # 정답 hwp파일 저장할거야
        time.sleep(0.2)
        print(f"{cnt_mizu}번 중 {i}번 정답 추출 완료")
        progress_condi.insert(END, f"{cnt_mizu}번 중 {i}번 정답 추출 완료\n")
        progress_condi.see(END)
        progress_condi.update()
        hwp.XHwpDocuments.Item(1).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)

def Divide_files(file_fullname):
    
    hwp.Open(file_fullname,"HWP","forceopen:True")
    time.sleep(2)

    count_mizu() # 미주 개수 세기 -> 변수 cnt_mizu 에 저장
    print(f"미주 개수는 {cnt_mizu}개 입니다.")
    progress_condi.insert(END, f"미주 개수는 {cnt_mizu}개 입니다.\n")
    progress_condi.see(END)
    progress_condi.update()

    preview_sol_hwp() # 빠른정답만들기(각파일로) (새탭이 없을때 사용해야함.)


    onepageoneproblem() # 한쪽에 한문제 (첫페이지는 표지)
    print("한쪽에 한문제 나누기 완료")
    progress_condi.insert(END, "한쪽에 한문제 나누기 완료\n")
    progress_condi.see(END)
    progress_condi.update()


    ### 문제 해설 나눠서 각 파일로 저장
    for i in range(1, cnt_mizu+1):
    # for i in range(1, 3): # 실험용
        tabdiv_pro_sol() # 원본 문제는 지우고 1탭에 문제만 2탭에 해설만 / 작동 후 해설 맨 처음
        save_sol_hml(i)
        save_sol_png(i)
        equation_to_text_first(hwp) #첫수식 text화
        save_sol_hwp(i)
        time.sleep(0.2)
        ## 해설 페이지가 1이면 이름만 바꾸고 2이상이면 병합할꺼야
        if solution_page==1:
            sol_png_name = file_fullname.replace(".hwp", "") + " "+ str(i).zfill(3) + "번 [2해설] [2png]"
            os.rename(sol_png_name+"001"+".png", sol_png_name+".png")
        else:
            image_merge()

        time.sleep(0.2)
        hwp.HAction.Run("WindowNextTab")
        hwp.HAction.Run("WindowNextTab")
        save_pro_hml(i)
        save_pro_png(i)
        time.sleep(0.2)
        hwp.HAction.Run("WindowNextTab")
        hwp.XHwpDocuments.Item(2).Close(isDirty=False)
        hwp.XHwpDocuments.Item(1).Close(isDirty=False)
        print(f"{cnt_mizu}번 중 {i}번 나누기 완료")
        progress_condi.insert(END, f"{cnt_mizu}번 중 {i}번 나누기 완료\n")
        progress_condi.see(END)
        progress_condi.update()


    hwp.XHwpDocuments.Item(0).Close(isDirty=False)
    # hwp.Quit()

    time.sleep(1)


    # 아래는 쪼개진 파일들 중 해설 비어있는거 찾아서 [빈해설] 붙이기 / 한글 파일 찾아서 ole객채?불러서 
    for num in range(1, cnt_mizu+1):
        sol_hml_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [1hml]"
        sol_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [2png]"
        sol_hwp_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [3hwp]"

        f = olefile.OleFileIO(sol_hwp_name+".hwp") # HWP 파일 열기
        encoded_text = f.openstream('PrvText').read() # PrvText 스트림의 내용 꺼내기
        decoded_text = encoded_text.decode('UTF-16') # 유니코드를 UTF-16으로 디코딩
        f.close()
        if len(decoded_text)==2: #decoded_text 이거 길이가 2이면 안에 내용 없는거임 
            os.rename(sol_hml_name+".hml", sol_hml_name+"[빈해설]"+".hml")
            os.rename(sol_png_name+".png", sol_png_name+"[빈해설]"+".png")
            os.remove(sol_hwp_name+".hwp")
        elif len(decoded_text)>2:
            os.remove(sol_hwp_name+".hwp")
        print(f"{cnt_mizu}번 중 {num}번 빈해설 확인 완료")
        progress_condi.insert(END, f"{cnt_mizu}번 중 {num}번 빈해설 확인 완료\n")
        progress_condi.see(END)
        progress_condi.update()

    time.sleep(3)
    # pyautogui.alert(text='작업이 끝났습니다.')
    






# 파일 추가
def add_file():
    files = filedialog.askopenfilenames(title="분할할 hwp파일을 선택하세요", \
        filetypes=(("HWP 파일", "*.hwp"), ("모든 파일", "*.*")), \
        initialdir=os.path.expanduser(r"~\Desktop"))
        # 최초에 사용자가 지정한 경로를 보여줌
    
    # 사용자가 선택한 파일 목록
    for file in files:
        list_file.insert(END, file)
        list_file.see(END)

# 선택 삭제
def del_file():
    #print(list_file.curselection())
    for index in reversed(list_file.curselection()):
        list_file.delete(index)


# 시작
def start():
    # 파일 목록 확인
    if list_file.size() == 0:
        msgbox.showwarning("경고", "분할할 hwp파일을 선택하세요")
        return
    else:
        result_div()
        msgbox.showinfo("알림", "작업이 완료되었습니다.")
        progress_head.delete(1.0, END)
        


def result_div(): #파일명 나누기 + 쪼개기
    file_fullnames = [os.path.split(x) for x in list_file.get(0, END)] # os.path.split -> 폴더, 파일명 나눠주는거 튜플로 나눠줌
    
    # print(file_fullnames)
    for idx, fullname in enumerate(file_fullnames) :
        dir = r'{0}'.format(fullname[0]).replace("/", "\\")
        name = fullname[1]
        name_only = re.sub(r".hwp", "", name)
        reidx = idx+1
        progress_head.delete(1.0, END)
        progress_head.insert(END, f"{reidx}번째 파일을 진행합니다.")
        progress_head.update()
        global file_fullname
        file_fullname = os.path.join(dir, name)
        print(file_fullname)
        # 파일 켜기
        Divide_files(file_fullname)
        progress_condi.delete(1.0, END)
        progress_condi.update()
        try: 
            if not os.path.exists(os.path.join(dir, name_only)): os.makedirs(os.path.join(dir, name_only)) 
        except OSError:
            print("Error: Cannot create the directory {}".format(os.path.join(dir, name_only)))
        time.sleep(1)
        lists = [i for i in os.listdir(dir) if i.startswith(f'{name_only}')] # name_only로 시작하는 파일 찾아라

        for list in lists: # 지금 작업하는 파일 이름(name_only)이 있는 파일 다 옮겨라 
            shutil.move(os.path.join(dir, list), os.path.join(dir, name_only))

        
        
        


        



# 파일 프레임 (파일 추가, 선택 삭제)
file_frame = Frame(root)
file_frame.pack(fill="x", padx=5, pady=5) # 간격 띄우기

btn_add_file = Button(file_frame, padx=5, pady=5, width=12, text="파일추가", command=add_file)
btn_add_file.pack(side="left")

btn_del_file = Button(file_frame, padx=5, pady=5, width=12, text="선택삭제", command=del_file)
btn_del_file.pack(side="right")

# 리스트 프레임
list_frame = LabelFrame(root, width=100, text="선택된 파일")
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=10, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)



# 진행 상태 확인
progress_frame = LabelFrame(root, width=100, height=5, text="진행상황")
progress_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(progress_frame)
scrollbar.pack(side="right", fill="y")

progress_head = Text(progress_frame, width=100, height=1)
progress_head.pack()

progress_condi = Text(progress_frame, width=100, height=30, yscrollcommand=scrollbar.set)
progress_condi.pack(side="left", fill="both", expand=True)
scrollbar.config(command=progress_condi.yview)


# progress_frame_file = Listbox(progress_frame, selectmode="extended", height=10, yscrollcommand=scrollbar.set)
# progress_frame_file.pack(side="left", fill="both", expand=True)
# scrollbar.config(command=list_file.yview)


# # 진행 상황 Progress Bar
# frame_progress = LabelFrame(root, text="진행상황")
# frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)

# p_var = DoubleVar()
# progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
# progress_bar.pack(fill="x", padx=5, pady=5)

# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)


# root.geometry('1000x1100')
# root.resizable(True, True)
root.resizable(False, False)
root.mainloop()



hwp.Quit()