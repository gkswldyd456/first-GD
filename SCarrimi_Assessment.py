import time
import pyautogui
import os, shutil, sys
import fnmatch 
import re

from bs4 import BeautifulSoup

from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from webdriver_manager.chrome import ChromeDriverManager





# def Select_GUGUN(s1, s2, s3): # 중/고 , 시 , 구군 까지 선택
    
#     driver.find_element_by_css_selector('#contents > div > div.m_choice_pbdata > div > div:nth-child(1) > div.input-group > ul > li:nth-child({0}) > label > span'.format(s1)).click() 
#     # s1번째 학교버전 (1=초 2=중 3=고 4=기타)
#     time.sleep(0.2)
    
    
#     html = driver.page_source
#     soup = BeautifulSoup(html)
#     items2 = soup.select('#mLevel2 > option')
    
#     global SDV
#     global SDN
    
#     SDV = []
#     SDN = []

#     for i in items2:
#         SDV.append(i["value"])

#     for i in items2:
#         SDN.append(i.text)

#     # del SDV[0] # 첫번째 이상한거 원소 삭제
#     # del SDN[0]
#     time.sleep(1)

    
#     driver.find_element_by_css_selector('#mLevel2 > option:nth-child({0})'.format(s2+1)).click() # s2번째 지역(시) (첫시1)
#     time.sleep(0.2)


#     html = driver.page_source
#     soup = BeautifulSoup(html)
#     items3 = soup.select('#mLevel3 > option')
    
#     global GGV
#     global GGN    
    
#     GGV = []
#     GGN = []

#     for i in items3:
#         GGV.append(i["value"])

#     for i in items3:
#         GGN.append(i.text)

#     # del GGV[0] # 첫번째 이상한거 원소 삭제
#     # del GGN[0]

#     time.sleep(1)


#     driver.find_element_by_css_selector('#mLevel3 > option:nth-child({0})'.format(s3+1)).click() # s3번째 지역(구) (첫구1)
#     time.sleep(0.2)


#     html = driver.page_source
#     soup = BeautifulSoup(html)
#     items4 = soup.select('#mLevel4 > option')
    
#     global SCV
#     global SCN
    
#     SCV = []
#     SCN = []

#     for i in items4:
#         SCV.append(i["value"])

#     for i in items4:
#         SCN.append(i.text)
        
#     # del SCV[0] # 첫번째 이상한거 원소 삭제
#     # del SCN[0]
    
#     time.sleep(1)

def Select_CLN(s1):
    
    driver.find_element_by_css_selector('#contents > div > div.m_choice_pbdata > div > div:nth-child(1) > div.input-group > ul > li:nth-child({0}) > label > span'.format(s1)).click() 
    # s1번째 학교버전 (1=초 2=중 3=고 4=기타)
    time.sleep(0.2)
    
    
    html = driver.page_source
    soup = BeautifulSoup(html)
    items2 = soup.select('#mLevel2 > option')
    
    global SDV
    global SDN
    
    SDN = []


    for i in items2:
        SDN.append(i.text)

    time.sleep(0.2)

def Select_SDN(s2):
    driver.find_element_by_css_selector('#mLevel2 > option:nth-child({0})'.format(s2+1)).click() # s2번째 지역(시) (첫시1)
    time.sleep(0.2)

    html = driver.page_source
    soup = BeautifulSoup(html)
    items3 = soup.select('#mLevel3 > option')
    
    global GGN    
    
    GGN = []

    for i in items3:
        GGN.append(i.text)

    time.sleep(0.2)

def Select_GGN(s3):
    driver.find_element_by_css_selector('#mLevel3 > option:nth-child({0})'.format(s3+1)).click() # s3번째 지역(구) (첫구1)
    time.sleep(0.2)

    html = driver.page_source
    soup = BeautifulSoup(html)
    items4 = soup.select('#mLevel4 > option')
    

    global SCV
    global SCN

    SCV = []
    SCN = []

    for i in items4:
        SCV.append(i["value"])

    for i in items4:
        SCN.append(i.text)
   
    time.sleep(0.2)



def Select_SCHOOL(s4): # 학교선택 후 검색누르기
    driver.find_element_by_css_selector('#mLevel4 > option:nth-child({0})'.format(s4+1)).click() # s4번째 학교 (첫학교1)
    time.sleep(0.5)

    driver.find_element_by_css_selector('#mobileSearchButton').click() # 검색버튼
    time.sleep(1)
    
    driver.switch_to.window(driver.window_handles[1]) #첫번째 탭으로 이동 (셀레니움 팝업창)
    time.sleep(0.5)

    # driver.find_element_by_css_selector('#gsYear > option:nth-child(2)').click() # 2021년 누르기
    # driver.find_element_by_css_selector('#gsYearBtn').click() # 선택버튼 누르기
    # time.sleep(0.5)

    # driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    # time.sleep(0.5)
    
    
def Assessment0(): # 학교선택 후 검색누르기

    driver.find_element_by_css_selector('#gsYear > option:nth-child(2)').click() # 2021년 누르기
    driver.find_element_by_css_selector('#gsYearBtn').click() # 선택버튼 누르기
    time.sleep(0.5)

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.5)


def Assessment1(): # 교과별평가계획 검색 누르기 까지
    driver.find_element_by_css_selector('body > div > div.publicInfo > div.pws_tabs_container.pws_tabs_horizontal.pws_tabs_horizontal_top.pws_tabs_noeffect.pws_none > ul > li:nth-child(6) > a').click() # 학업성취 사항 체크
    time.sleep(0.2)

    driver.find_element_by_css_selector('body > div > div.publicInfo > div.pws_tabs_container.pws_tabs_horizontal.pws_tabs_horizontal_top.pws_tabs_noeffect.pws_none > div > div.pws_tab_single.pws_hide.pws_show > ul > li:nth-child(1) > a').click() # 교과별 학업성취 사항 체크
    time.sleep(1)

    html = driver.page_source
    soup = BeautifulSoup(html)
    GSNY = soup.select('#select_trans_dt > option')
    
    global GSNYV
    global GSNYN
    
    GSNYV = []
    GSNYN = []

    for i in GSNY:
        GSNYV.append(i["value"])

    for i in GSNY:
        GSNYN.append(i.text)

    time.sleep(0.5)


def Cheking(): # 해당 파일 부분 리스트로 정보 받기
    html = driver.page_source
    soup = BeautifulSoup(html)
    FileN1 = soup.select('#gongsiInfo > div.table_wrap > div:nth-child(3) > div.attached_file > a')
    FileN2 = soup.select('#gongsiInfo > div.table_wrap > div:nth-child(4) > div.attached_file > a')
    FileN3 = soup.select('#gongsiInfo > div.table_wrap > div > a')



    global FileN1_1 # 1교과별(학년별)평가계획에 관한 사항 있는곳
    global FileN2_1 # 2학업성적관리 규정 있는곳
    global FileN3_1 # 아무것도 안써있는곳

    FileN1_1 = []
    FileN2_1 = []
    FileN3_1 = []

    for i in FileN1:
        FileN1_1.append(re.sub('\(\d*(,\d\d\d)*\s*KB\)', "", i.text))#뒤에 KB부분 없앤 이름만 리스트로 넣기

    for i in FileN2:
        FileN2_1.append(re.sub('\(\d*(,\d\d\d)*\s*KB\)', "", i.text))
        
    for i in FileN3:
        FileN3_1.append(re.sub('\(\d*(,\d\d\d)*\s*KB\)', "", i.text))
    
    time.sleep(0.5)
    

def Assessment2(s4): # 
    path1 = os.path.expanduser('~\\Downloads\\') # 다운로드경로
    path2 = os.path.expanduser('~\\downloads\\') # 다운로드경로
    global path
    
    if os.path.exists(path1)==True:
        path = path1
    elif os.path.exists(path2)==True:
        path = path2
    else:
        pyautogui.alert(text='Downloads 되는 폴더이름을 담당자에게 알려주세요.')
    
    global Dname
    Dname = "["+str(sel1) +str(sel2).zfill(2)+str(sel3).zfill(2)+str(s4).zfill(2)+"]"+str(SCN[s4])+"["+str(SCV[s4])+"]" # 파일 이름 중 : 학교이름+[학교알리미코드]

    try: #학교이름+알리미코드 이름의 폴더 만들어
        if not os.path.exists(path+Dname): os.makedirs(path+Dname) 

    except OSError:
        print("Error: Cannot create the directory {}".format(path+Dname))
    
    time.sleep(1)
    



def Assessment3():
    FileN1_1N = " [1]교과별 평가계획"
    FileN2_1N = " [2]학업성적관리규정"
    FileN3_1N = " [3]평가계획and성적관리규정"
    
    if len(FileN1_1) != 0 : 
        try: 
            if not os.path.exists(path+Dname+str(GSNYNI)+FileN1_1N): os.makedirs(path+Dname+str(GSNYNI)+FileN1_1N) 

        except OSError:
            print("Error: Cannot create the directory {}".format(path+Dname+str(GSNYNI)+FileN1_1N))
        
        time.sleep(1)
        
        if len(FileN1_1) == 1 : 
            driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div:nth-child(3) > div.attached_file > a').click()
            time.sleep(0.5)
            while os.path.isfile(path+Dname+str(GSNYNI)+FileN1_1N+"\\"+str(FileN1_1[0]))==False:
                try:
                    shutil.move(path+str(FileN1_1[0]), path+Dname+str(GSNYNI)+FileN1_1N)
                    time.sleep(0.5)
                except FileNotFoundError:
                    pass
                except PermissionError:
                    pass
                
        else :
            for i in range(1,len(FileN1_1)+1):
                driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div:nth-child(3) > div.attached_file > a:nth-child({0})'.format(i)).click()
                time.sleep(0.5)
                reFileN1_1 = re.sub('&', "&amp;", FileN1_1[i-1])
                while os.path.isfile(path+Dname+str(GSNYNI)+FileN1_1N+"\\"+str(reFileN1_1))==False:
                    try:
                        shutil.move(path+str(reFileN1_1), path+Dname+str(GSNYNI)+FileN1_1N)
                        time.sleep(0.5)
                    except FileNotFoundError:
                        pass
                    except PermissionError:
                        pass
        shutil.move(path+Dname+str(GSNYNI)+FileN1_1N, path+Dname)
        time.sleep(0.5)
        # pyautogui.alert()
        
    else : 
        pass
    
    if len(FileN2_1) != 0 : 
        try: 
            if not os.path.exists(path+Dname+str(GSNYNI)+FileN2_1N): os.makedirs(path+Dname+str(GSNYNI)+FileN2_1N) 

        except OSError:
            print("Error: Cannot create the directory {}".format(path+Dname+str(GSNYNI)+FileN2_1N))
        
        time.sleep(1)
        
        if len(FileN2_1) == 1 : 
            driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div:nth-child(4) > div.attached_file > a').click()
            time.sleep(0.5)
            while os.path.isfile(path+Dname+str(GSNYNI)+FileN2_1N+"\\"+str(FileN2_1[0]))==False:
                try:
                    shutil.move(path+str(FileN2_1[0]), path+Dname+str(GSNYNI)+FileN2_1N)
                    time.sleep(0.5)
                except FileNotFoundError:
                    pass
                except PermissionError:
                    pass
        else :
            for i in range(1,len(FileN2_1)+1):
                driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div:nth-child(4) > div.attached_file > a:nth-child({0})'.format(i)).click()
                time.sleep(0.5)
                reFileN2_1 = re.sub('&', "&amp;", FileN2_1[i-1])
                while os.path.isfile(path+Dname+str(GSNYNI)+FileN2_1N+"\\"+str(reFileN2_1))==False:
                    try:
                        shutil.move(path+str(reFileN2_1), path+Dname+str(GSNYNI)+FileN2_1N)
                        time.sleep(0.5)
                    except FileNotFoundError:
                        pass
                    except PermissionError:
                        pass
        shutil.move(path+Dname+str(GSNYNI)+FileN2_1N, path+Dname)
        time.sleep(0.5)
        # pyautogui.alert()
        
    else : 
        pass
    
    if len(FileN3_1) != 0 : 
        try: 
            if not os.path.exists(path+Dname+str(GSNYNI)+FileN3_1N): os.makedirs(path+Dname+str(GSNYNI)+FileN3_1N) 

        except OSError:
            print("Error: Cannot create the directory {}".format(path+Dname+str(GSNYNI)+FileN3_1N))
        
        time.sleep(1)
        
        if len(FileN3_1) == 1 : 
            driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div > a').click()
            time.sleep(0.5)
            while os.path.isfile(path+Dname+str(GSNYNI)+FileN3_1N+"\\"+str(FileN3_1[0]))==False:
                try:
                    shutil.move(path+str(FileN3_1[0]), path+Dname+str(GSNYNI)+FileN3_1N)
                    time.sleep(0.5)
                except FileNotFoundError:
                    pass
                except PermissionError:
                    pass
        else :
            for i in range(1,len(FileN3_1)+1):
                driver.find_element_by_css_selector('#gongsiInfo > div.table_wrap > div > a:nth-child({0})'.format(i)).click()
                time.sleep(0.5)
                reFileN3_1 = re.sub('&', "&amp;", FileN3_1[i-1])
                while os.path.isfile(path+Dname+str(GSNYNI)+FileN3_1N+"\\"+str(reFileN3_1))==False:
                    try:
                        shutil.move(path+str(reFileN3_1), path+Dname+str(GSNYNI)+FileN3_1N)
                        time.sleep(0.5)
                    except FileNotFoundError:
                        pass
                    except PermissionError:
                        pass
        shutil.move(path+Dname+str(GSNYNI)+FileN3_1N, path+Dname)
        time.sleep(0.5)
        # pyautogui.alert()
        
    else : 
        pass
    
    time.sleep(1)
    # pyautogui.alert(text='다음 공시년월이 있다면 넘어갑니다.')


def ActionONESC(s4): # 한 학교만 하기
    Select_SCHOOL(s4)
    pyautogui.moveRel(0, -10, 0.1)
    try:   
        (driver.find_element_by_css_selector('body > div > div.basicInfo > div.basic_data > p > span.closed_school').is_enabled()==True or  
        driver.find_element_by_css_selector('body > div > div > div > p > span.closed_school').is_enabled()==True) # 폐교확인
        print("\n\n폐교입니다.\n\n")
    except NoSuchElementException: 
        Assessment0()
        Assessment1()
        Assessment2(s4)
        for i in range(1,len(GSNYN)+1):
            driver.find_element_by_css_selector('#select_trans_dt > option:nth-child({0})'.format(i)).click() # 자료제출일 선택
            time.sleep(1)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            Cheking()
            global GSNYNI
            GSNYNI = GSNYN[i-1]
            Assessment3()
        time.sleep(1)
        # pyautogui.alert(text='다음 학교가 있다면 다음 학교로 넘어갑니다.')
        # time.sleep(1)
    
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    driver.execute_script('window.scrollTo(0,0)') # 페이지 맨 위로 가라
    time.sleep(1)
    pyautogui.moveRel(0, 10, 0.1)


def ActionSCS(): # 학교 전체 다 돌리기
    for num in range(1, len(SCN)):
        Select_SCHOOL(num)
        pyautogui.moveRel(0, -10, 0.1)
        try:   
            (driver.find_element_by_css_selector('body > div > div.basicInfo > div.basic_data > p > span.closed_school').is_enabled()==True or  
            driver.find_element_by_css_selector('body > div > div > div > p > span.closed_school').is_enabled()==True) # 폐교확인
            print("\n\n폐교입니다.\n\n")
        except NoSuchElementException: 
            Assessment0()
            Assessment1()
            Assessment2(num)
            for i in range(1,len(GSNYN)+1):
                driver.find_element_by_css_selector('#select_trans_dt > option:nth-child({0})'.format(i)).click() # 자료제출일 선택
                time.sleep(1)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                Cheking()
                global GSNYNI
                GSNYNI = GSNYN[i-1]
                Assessment3()
            time.sleep(1)
            # pyautogui.alert(text='다음 학교가 있다면 다음 학교로 넘어갑니다.')
            # time.sleep(1)
        
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)
        driver.execute_script('window.scrollTo(0,0)') # 페이지 맨 위로 가라
        time.sleep(1)
        pyautogui.moveRel(0, 10, 0.1)
        
    


########## 아래부터 실제 사용

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()

driver.maximize_window() # 윈도우창 최대
url = 'https://www.schoolinfo.go.kr/ei/ss/pneiss_a03_s0.do' # 학교알리미 학교별 공시정보
driver.get(url) # 페이지 열어라 
time.sleep(1)


pyautogui.keyDown('win') # win 키를 누른 상태를 유지합니다.
pyautogui.press('left') # 왼화살표를 입력합니다. 
pyautogui.keyUp('win') # win 키를 뗍니다.  -> 화면 왼쪽절반차지
time.sleep(0.2)

size = driver.get_window_size()
height = size.get("height")
driver.set_window_size(869, int(height))
time.sleep(1)






##### 아래 주석 조정하면서 돌려

# pyautogui.alert(text='[중/고 -> 시/도 -> 구/군 -> 학교]을 선택하면 한 학교가 작업됩니다.')

global sel1
global sel2
global sel3

# ver1 한 학교만 돌려

# sel1 = int(pyautogui.prompt(text='중/고' + ' <- 중/고 숫자 쓰세요'))
# sel2 = int(pyautogui.prompt(text='시/도' + ' <- 시/도 숫자 쓰세요'))
# sel3 = int(pyautogui.prompt(text='구/군' + ' <- 구/군 숫자 쓰세요'))
# num = int(pyautogui.prompt(text='학교' + ' <- 학교 숫자 쓰세요'))

# Select_CLN(sel1)
# Select_SDN(sel2)
# Select_GGN(sel3)
# ActionONESC(num)


# ver2 다 돌려 

# for i in range(2,4): #range(1,len(GBN)):
#     Select_CLN(i)
#     sel1 = i
#     for j in range(1,len(SDN)):
#         Select_SDN(j)
#         sel2 = j
#         for k in range(1,len(GGN)):
#             Select_GGN(k)
#             sel3 = k
#             ActionSCS()



# ver3-1 특정학교 부터 돌려

sel1 = int(pyautogui.prompt(text='중/고' + ' <- 중/고 숫자 쓰세요'))
sel2 = int(pyautogui.prompt(text='시/도' + ' <- 시/도 숫자 쓰세요'))
sel3 = int(pyautogui.prompt(text='구/군' + ' <- 구/군 숫자 쓰세요'))

Select_CLN(sel1)
Select_SDN(sel2)
Select_GGN(sel3)
for num in range(12,len(SCN)): # 여기 수정해 -> 어떤 학교부터
    ActionONESC(num)



# ver3-2 특정구 부터 돌리기

# sel1 = int(pyautogui.prompt(text='중/고' + ' <- 중/고 숫자 쓰세요'))
# sel2 = int(pyautogui.prompt(text='시/도' + ' <- 시/도 숫자 쓰세요'))

# Select_CLN(sel1)
# Select_SDN(sel2)
# for k in range(1,len(GGN)): # 여기 수정해 -> 어떤 구부터
#     Select_GGN(k)
#     sel3 = k
#     ActionSCS()



# # # ver3-3 특정지역(시) 부터 돌리기

# sel1 = int(pyautogui.prompt(text='중/고' + ' <- 중/고 숫자 쓰세요'))

# Select_CLN(sel1)
# for j in range(2,len(SDN)): # 여기 수정해
#     Select_SDN(j)
#     sel2 = j
#     for k in range(1,len(GGN)):
#         Select_GGN(k)
#         sel3 = k
#         ActionSCS()

    


########2 01 12 부터 시작하면 됨~~~


pyautogui.alert(text='작업이 끝났습니다.')









