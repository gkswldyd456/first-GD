import time
import pyautogui
import os, shutil, sys
import fnmatch 

from bs4 import BeautifulSoup

from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from webdriver_manager.chrome import ChromeDriverManager





def Select_GUGUN(s1, s2, s3): # 중/고 , 시 , 구군 까지 선택
    
    driver.find_element_by_css_selector('#contents > div > div.m_choice_pbdata > div > div:nth-child(1) > div.input-group > ul > li:nth-child({0}) > label > span'.format(s1)).click() 
    # s1번째 학교버전 (1=초 2=중 3=고 4=기타)
    time.sleep(0.2)
    
    
    html = driver.page_source
    soup = BeautifulSoup(html)
    items2 = soup.select('#mLevel2 > option')
    
    global SDV
    global SDN
    
    SDV = []
    SDN = []

    for i in items2:
        SDV.append(i["value"])

    for i in items2:
        SDN.append(i.text)

    # del SDV[0] # 첫번째 이상한거 원소 삭제
    # del SDN[0]
    time.sleep(1)

    
    driver.find_element_by_css_selector('#mLevel2 > option:nth-child({0})'.format(s2+1)).click() # s2번째 지역(시) (첫시1)
    time.sleep(0.2)


    html = driver.page_source
    soup = BeautifulSoup(html)
    items3 = soup.select('#mLevel3 > option')
    
    global GGV
    global GGN    
    
    GGV = []
    GGN = []

    for i in items3:
        GGV.append(i["value"])

    for i in items3:
        GGN.append(i.text)

    # del GGV[0] # 첫번째 이상한거 원소 삭제
    # del GGN[0]

    time.sleep(1)


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
        
    # del SCV[0] # 첫번째 이상한거 원소 삭제
    # del SCN[0]
    
    time.sleep(1)



def Select_SCHOOL(s4): # 학교선택 후 검색누르기
    driver.find_element_by_css_selector('#mLevel4 > option:nth-child({0})'.format(s4+1)).click() # s4번째 학교 (첫학교1)
    time.sleep(0.2)

    driver.find_element_by_css_selector('#mobileSearchButton').click() # 검색버튼
    time.sleep(1)



def Achievement1(): # 교과별학업성취 검색 누르기 까지
    driver.find_element_by_css_selector('body > div > div.publicInfo > div.pws_tabs_container.pws_tabs_horizontal.pws_tabs_horizontal_top.pws_tabs_noeffect.pws_none > ul > li:nth-child(6) > a').click() # 학업성취 사항 체크
    time.sleep(0.2)

    driver.find_element_by_css_selector('body > div > div.publicInfo > div.pws_tabs_container.pws_tabs_horizontal.pws_tabs_horizontal_top.pws_tabs_noeffect.pws_none > div > div.pws_tab_single.pws_hide.pws_show > ul > li:nth-child(2) > a').click() # 교과별 학업성취 사항 체크
    time.sleep(1)



def Achievement2(): # 검색 누른 후 자동입력방지 실행

    while True : # 자동방지 뜨는거 입력받아서 만약 잘못써도 다시 실행하게
        try: 
            driver.find_element_by_css_selector('#catpcha44 > img').is_enabled()==True
            el = driver.find_element_by_css_selector('#passLine44')
            el.click()
            time.sleep(0.2)
            btn_1 = pyautogui.prompt(text='자동입력방지입력')
            el.send_keys(btn_1)
            # el.send_keys(input('자동입력방지입력 : '))
            el.send_keys(Keys.ENTER)
            time.sleep(1)
        except NoSuchElementException:
            break
    time.sleep(1.5)


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

    time.sleep(1)



def Achievement3(s4): # 엑셀다운로드 반복 [공시년월 1개 있을때 / 1개 이상 있을 떄 / 없을때]
    path1 = os.path.expanduser('~\\Downloads\\') # 다운로드경로
    path2 = os.path.expanduser('~\\downloads\\') # 다운로드경로
    
    if os.path.exists(path1)==True:
        path = path1
    elif os.path.exists(path2)==True:
        path = path2
    else:
        pyautogui.alert(text='Downloads 되는 폴더이름을 담당자에게 알려주세요.')
    
    Dname = str(SCN[s4])+"["+str(SCV[s4])+"]" # 파일 이름 중 : 학교이름+[학교알리미코드]
    
    
    if len(GSNYN)==1:
        driver.find_element_by_css_selector('#gongsiInfo > div.tabtt_wrap > div > div > a:nth-child(6)').click() # 엑셀다운로드 누르기
        time.sleep(1)
                                            
        os.rename(path + "교과별 학업성취 사항.xls", path + Dname +str(GSNYN[0])+".xls") # 학교이름+알리미코드+공시년월내용
        time.sleep(1)
        
        try: #학교이름+알리미코드 이름의 폴더 만들어
            if not os.path.exists(path+Dname): os.makedirs(path+Dname) 

        except OSError:
            print("Error: Cannot create the directory {}".format(path+Dname))


        for f in fnmatch.filter(os.listdir(path),'*.xls'): # 엑셀다운받은것들 위 폴더로 옮겨
            if Dname in f:
                file_to_move = os.path.join(path, f)
                shutil.move(file_to_move, path+Dname)
        time.sleep(1)
    
    
    elif len(GSNYN)>1:
        driver.find_element_by_css_selector('#gongsiInfo > div.tabtt_wrap > div > div > a:nth-child(6)').click() # 엑셀다운로드 누르기
        time.sleep(1)

        os.rename(path + "교과별 학업성취 사항.xls", path + Dname +str(GSNYN[0])+".xls")
        time.sleep(1)
        for i in range(2,len(GSNYN)+1):
            driver.find_element_by_css_selector('#select_trans_dt > option:nth-child({0})'.format(i)).click()
            time.sleep(1) 
            while True : # 자동방지 뜨는거 입력받아서 만약 잘못써도 다시 실행하게
                try: 
                    driver.find_element_by_css_selector('#catpcha44 > img').is_enabled()==True
                    el = driver.find_element_by_css_selector('#passLine44')
                    el.click()
                    time.sleep(0.2)
                    btn_1 = pyautogui.prompt(text='자동입력방지입력')
                    el.send_keys(btn_1)
                    # el.send_keys(input('자동입력방지입력 : '))
                    el.send_keys(Keys.ENTER)
                    time.sleep(1)
                except NoSuchElementException:
                    break
            time.sleep(1.5)
            driver.find_element_by_css_selector('#gongsiInfo > div.tabtt_wrap > div > div > a:nth-child(6)').click() # 엑셀다운로드 누르기
            time.sleep(1)
            GSNYNI = GSNYN[i-1]
            os.rename(path + "교과별 학업성취 사항.xls", path + Dname +str(GSNYNI)+".xls")
            time.sleep(1)
            
        try: 
            if not os.path.exists(path+Dname): os.makedirs(path+Dname) 

        except OSError:
            print("Error: Cannot create the directory {}".format(path+Dname))

        for f in fnmatch.filter(os.listdir(path),'*.xls'):
            if Dname in f:
                file_to_move = os.path.join(path, f)
                shutil.move(file_to_move, path+Dname)

        time.sleep(1)

    else:
        pass





########## 아래부터 실제 사용

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()


# driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window() # 윈도우창 최대
# driver.implicitly_wait(10)
# actions = ActionChains(driver)
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





# 학교선택 전까지 한 후 학교선택 및 뒷 작업 반복

pyautogui.alert(text='[중/고 -> 시/도 -> 구/군]을 선택하면 학교 전체가 작업됩니다.')

a = int(pyautogui.prompt(text='중/고' + ' <- 중/고 숫자 쓰세요'))
b = int(pyautogui.prompt(text='시/도' + ' <- 시/도 숫자 쓰세요'))
c = int(pyautogui.prompt(text='구/군' + ' <- 구/군 숫자 쓰세요'))

Select_GUGUN(a, b, c)
for num in range(1, len(SCN)):
    Select_SCHOOL(num)
    driver.switch_to.window(driver.window_handles[1]) #첫번째 탭으로 이동
    time.sleep(0.2)

    driver.find_element_by_css_selector('#gsYear > option:nth-child(2)').click()
    driver.find_element_by_css_selector('#gsYearBtn').click()
    time.sleep(0.2)

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    Achievement1()
    try: 
        driver.find_element_by_css_selector('#catpcha44 > img').is_enabled()==True # 캡챠그림 있는지 확인하고 있으면 해라
        Achievement2()
        Achievement3(num)
        # driver.execute_script('window.scrollTo(0,0)') # 페이지 맨 위로 가라
        time.sleep(1)
    except NoSuchElementException: # 예외나오면 스크롤 맨위로 올리고 1초기다려(안하니 작업자 생각상 이상)
        driver.execute_script('window.scrollTo(0,0)')
        time.sleep(1)
        pyautogui.alert(text='['+SCN[num]+'] '+'캡차가 안잡혀서 예외 처리가 되었습니다.')
        # input('캡차가 안잡혀서 예외 처리가 되었습니다.')
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(0.1)
    driver.execute_script('window.scrollTo(0,0)') # 페이지 맨 위로 가라
    time.sleep(1)
time.sleep(3)
pyautogui.alert(text='작업이 끝났습니다.')
driver.close()












