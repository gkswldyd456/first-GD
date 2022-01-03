import win32com.client as win
import os
import fnmatch
import time


# o=win.Dispatch("HWPFrame.HwpObject")
# o=win.gencache.EnsureDispatch
# o=win.dynamic.Dispatch("HWPFrame.HwpObject")



def Insert(dir,frame,name):
	os.rename(dir+name, dir+"working.hml")
	#파일명 길면 오류나는 것 같아서 짧게 순번으로 rename
	o=win.gencache.EnsureDispatch("HWPFrame.HwpObject")
	o.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
	#한글 권한 설정 보안 레지스트리
	o.XHwpWindows.Item(0).Visible=False
	#백그라운드에서 동작(False)
	o.Open(dir+frame,"HWP","forceopen:True")
	o.HAction.GetDefault("InsertFile", o.HParameterSet.HInsertFile.HSet); 
	option=o.HParameterSet.HInsertFile
	option.filename = dir+"working.hml"
	option.KeepSection = 0
	option.KeepCharshape = 0
	option.KeepParashape = 0
	option.KeepStyle = 0
	o.HAction.Execute("InsertFile", o.HParameterSet.HInsertFile.HSet)
	time.sleep(0.1)
	#오류 방지 딜레이 타임
	while os.path.isfile(dir+name.replace(".hml",".hwp"))!=True:
	#파일 저장 안되었을 시 반복	
		o.HAction.GetDefault("FileSaveAs_S", o.HParameterSet.HFileOpenSave.HSet)
		o.HParameterSet.HFileOpenSave.filename = dir+name.replace(".hml",".hwp")
		o.HParameterSet.HFileOpenSave.Format = "HWP"
		o.HAction.Execute("FileSaveAs_S", o.HParameterSet.HFileOpenSave.HSet)
	o.XHwpDocuments.Close(isDirty=False) # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
	o.Quit()
	os.remove(dir+"working.hml")

dir1 = "C:\\Users\\HanJiYong\\Desktop\\change\\"
#hml과 frame이 있는 폴더 경로
frame1 = "수학비서 타이핑 양식(수정)(2단)(필드적용).hwp"
#frame
names = fnmatch.filter(os.listdir(dir1),'*.hml')
#hml 파일만 리스트화

for i in names:
	Insert(dir1,frame1,i)




