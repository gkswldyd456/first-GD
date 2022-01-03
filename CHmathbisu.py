import win32com.client as win
import os
import fnmatch
import time
from PIL import Image


'''
InsertPicture
sizeoption 파라미터 설명
0 : 이미지 원래의 크기로 삽입한다. width와 height를 지정할 필요 없다.
1 : width와 height에 지정한 크기로 그림을 삽입한다.
2 : 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞게 자동 조절하여 삽입한다.
    width는 셀의 width만큼, height는 셀의 height만큼 확대/축소된다.
    캐럿이 셀 안에 있지 않으면 이미지의 원래 크기대로 삽입된다.
3 : 현재 캐럿이 표의 셀 안에 있을 경우, 셀의 크기에 맞추어 원본 이미지의 가로 세로의 비율이 동일하게 확대/축소하여 삽입한다.

imgae 패키지 resize로 하니까 화질이 너무 깨져버리네...sizeoption=3 으로하니까 화질 좋네
'''


'''
# field_list = [i for i in o.GetFieldList().split("\x02")]
# print('\n', field_list) #필드이름 호출

# o.MoveToField("시작왼") #시작왼 필드로 이동

# o.PutFieldText("시작왼", "ㅎㅎ") #시작왼 필드에 ㅎㅎ 글씨 써
'''



#수학비서기본틀 바탕쪽 이미지 수정하는 함수 (파일 킨 후 사용)
def CHimgbatang(dir,img): #CHimgbatang(경로,그림이름) 
    o.MoveToField("수학비서짝") #수학비서짝 필드로 이동
    o.HAction.Run("MoveSelRight")
    o.HAction.Run("DeleteBack")
    o.InsertPicture(dir+img, Embedded=True, sizeoption=3) #이미지 그냥 비율에 맞춰 넣기
    o.MoveToField("수학비서홀") #수학비서홀 필드로 이동
    o.HAction.Run("MoveSelRight")
    o.HAction.Run("DeleteBack")
    o.InsertPicture(dir+img, Embedded=True, sizeoption=3)

def CHimgLeft(dir,img): #CHimgLeft(경로,그림이름) 
    o.MoveToField("시작왼") #수학왼 필드로 이동
    o.InsertPicture(dir+img, Embedded=True, sizeoption=3) #이미지 그냥 비율에 맞춰 넣기

def CHimgRight(dir,img): #CHimgRight(경로,그림이름) 
    o.MoveToField("시작오") #수학오 필드로 이동
    o.InsertPicture(dir+img, Embedded=True, sizeoption=3) #이미지 그냥 비율에 맞춰 넣기

def CHmdname(name='타이핑 서비스'): #CHmdname(가운데제목이름 : 기본값 타이핑서비스) 
    o.MoveToField("시작이름")
    o.HAction.Run("SelectAll")
    o.HAction.Run("DeleteBack");
    o.PutFieldText("시작이름", name) #시작왼 필드에 name 글씨 써


dir1 = "C:\\Users\\HanJiYong\\Desktop\\change\\"
#hml과 frame이 있는 폴더 경로
frame1 = "수학비서 타이핑 양식(수정)(2단) - 복사본.hwp"
#frame
img1 = "name1.png" #실험용png
img2 = "name2.png" #실험용png
img3 = "name3.png" #실험용png



# o=win.Dispatch("HWPFrame.HwpObject") #이거 파라미터같은거 정확히써야해서 오류 많음
o=win.gencache.EnsureDispatch("HWPFrame.HwpObject")
o.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
#한글 권한 설정 보안 레지스트리
o.XHwpWindows.Item(0).Visible=True
#백그라운드에서 동작(False) / 보이게 동작(True)
o.Open(dir1+frame1,"HWP","forceopen:True")
CHimgbatang(dir1,img1)
CHimgLeft(dir1,img2)
CHimgRight(dir1,img3)
CHmdname("바꿔보자바꿔ㅇㅁㄴㅇㄴㅁㅇㄴㅁ바꿔")
# o.XHwpDocuments.Close(isDirty=False) # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
# o.Quit()



