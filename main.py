import pandas as pd #pip install pandas
import pyautogui as pg   #pip install pyautogui
import time
import keyboard   #pip install keyboard
import clipboard  #pip install clipboard
from PIL import ImageGrab
import pytesseract #pip install pytesseract #pip install openpyxl

#file -> open -> ilgum -> ok클릭 -> "NEW WINDOW" 로 열기

###프로그램 돌리기 전 확인사항###
# 1. chest PA 결과 이상소견 있는 사람들 기타 소견 입력하기
# 2. 검등2018 전체화면으로 만들기
# 3. 왼쪽 상단 사업장 체크표시 해제, 오른쪽 상단 일반에만 체크, 일자별 오늘 일자 선택, 원내/출장 선택, 검색 누르기
# 4. 검진종목 클릭해서 일반검진만으로 정렬
# 5. 저장후 자동 다음사람 선택 체크하기
# 6. mode는 2, rep은 오른쪽 하단 일반-특수 값 만큼, pjdate는 판정일(오늘날짜)

mode = 2   # mode1 : 1개씩, mode2 : 여러개 한꺼번에, mode3 : 생활습관만
            # mode 1일 경우 alt + s 누르면 한사람씩 돌아감
            # mode 2일 경우 esc 누르면 그 사람 판정 끝내고 중지
rep = 90    #반복횟수
pjdate ='20211227'   #판정일
code = "일검" #일검, 종검, 야간, 물질

###프로그램돌리고 나서 확인사항###
# 1. ilgumdf 파일을 이름 순으로 정렬해서 출력하고 다른이름으로 저장, ilgumdf파일은 끄기
# 2. 판정여부가 완료가 아닌 사람들은 직접 확인해줘야함
# 3. 오류가 났을 때
      # 값 입력이 정상적이지 않을 때(숫자, 문자, 특수문자), 검등2018 자체의 문제일때는 오류가 날 수 밖에 없으니 오류 난 사람은 직접 판정 넣어야함
      # 이외 오류가 날 상황은 많으니 우승희(010-8455-7735)로 알려주면 됨

dateLee = [1201, 1203, 1206, 1210, 1213, 1215, 1217, 1218, 1220, 1224, 1227, 1229, 1231,
    1101, 1103, 1105, 1108, 1112, 1115, 1117, 1119, 1122, 1126, 1129, 1201, 1203,
    802, 806, 809, 811, 813, 816, 820, 823, 825, 827, 828, 830,
    901, 903, 906, 908, 910, 911, 913, 916, 917, 924, 927,
    1001, 1004, 1006, 1008, 1013, 1015, 1018, 1021, 1025, 1027, 1029]
dateSong = [1202, 1207, 1208, 1209, 1214, 1216, 1221, 1222, 1223, 1228, 1230,
    1102, 1104, 1109, 1110, 1111, 1113, 1116, 1118, 1123, 1124, 1125, 1127, 1130, 1202,
    803, 804, 805, 810, 812, 817, 818, 819, 824, 826, 831, 925,
    902, 907, 909, 914, 915, 923, 928, 929, 930,
    1005, 1007, 1012, 1014, 1019, 1020, 1022, 1023, 1026, 1028]


datemode = 1
profmode = 1




if mode == 1 or mode == 2 :
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    exportdf = pd.DataFrame(columns=['사업장','검진일','성함','판정여부','정상A','정상B','일반질환','만성의심','유질환자','생활습관'])
    importujilwhan = pd.read_excel('C:\\Users\\DELL3040\\Desktop\\ujilwhan.xlsx')
    importgitar = pd.read_excel('C:\\Users\\DELL3040\\Desktop\\gitar.xlsx')

    ERROR = 0
    savecode = ""
    gitar1 = 0
    gitar2 = 0
    gitar3 = 0
    gitar4 = ""
    swresult =''

    ############################################################################################
    def smoking_chu():
        global swresult
        pg.click(x=244,y=550);pg.hotkey('ctrl','c');s_1=clipboard.paste();s_1=int(s_1)
        pg.click(x=294,y=600)
        if s_1<=3 :
            nicotine=1
            pg.press('1')
            #print("니코틴 의존도 낮음")
            swresult = swresult + '낮음 1     '
        elif s_1<=6 :
            nicotine=2
            pg.press('2')
            #print("니코틴 의존도 중간")
            swresult = swresult + '중간 1     '
        else :
            nicotine=3
            pg.press('3')
            #print("니코틴 의존도 높음")
            swresult = swresult + '높음 1     '
        pg.click(x=296,y=355);pg.hotkey('ctrl','c');s_2=clipboard.paste();s_2=int(s_2)
        if s_2==1 : cessation=3
        elif s_2==2 : cessation=2
        else : cessation=1

        pg.click(x=294,y=620);    pg.press('1')

        return nicotine,cessation

    def smoking_sangwhal(nicotine,cessation,screen_sang_1):

        if screen_sang_1.getpixel((45,262)) != (0, 0, 0):
            pg.click(45,262)
            time.sleep(0.1)

        pg.click(x=251,y=284);pg.press('2')

        if nicotine==1 : nicotine=1; pg.press('1')
        elif nicotine==2 : nicotine=2; pg.press('2')
        else : pg.press('3')

        if cessation==1 : nicotine=1; pg.press('1')
        elif cessation==2 : nicotine=2; pg.press('2')
        else : pg.press('3')

        pg.click(x=71,y=392)
        pg.click(x=75,y=623)
        pg.click(x=75,y=732)

    def drinking_chu():
        global swresult
        pg.click(x=580,y=693);pg.hotkey('ctrl','c');d_1=clipboard.paste();d_1=int(d_1)
        if d_1==2:
            #print("적정음주", end=' ')
            swresult = swresult + '적정음주 '
        elif d_1==3:
            #print("위험음주", end=' ')
            swresult = swresult + '위험음주 '
        else :
            #print("알코올 사용장애", end=' ')
            swresult = swresult + '알코올 사용장애 '

        #print("1")
        swresult = swresult + '1     '


        pg.click(x=580,y=717);pg.press('1')
        return d_1

    def drinking_sangwhal(d_1,screen_sang_1,r_1,r_2,r_3,r_4):

        if screen_sang_1.getpixel((355,263)) != (0, 0, 0):
            pg.click(355, 263)
            time.sleep(0.1)

        pg.click(x=381,y=356)
        if r_1==1 : pg.click(x=381,y=381)
        if r_2==1 : pg.click(x=381,y=405)
        if r_3==1 : pg.click(x=381,y=432)

        pg.click(x=518,y=292);
        if d_1==3 :
            pg.press('2')
            pg.click(x=389,y=654)
        elif d_1==4 :
            pg.press('3')
            pg.click(x=389,y=674)

    def exercise_chu(screen):
        global swresult
        e_1_1=0;e_1_h=0;e_1_m=0; e_2_1=0;e_2_h=0;e_2_m=0;
        pg.click(x=892, y=483);pg.hotkey('ctrl', 'c');e_1 = clipboard.paste();e_1 = int(e_1)
        if e_1==1:
            pg.click(x=892, y=504);pg.hotkey('ctrl', 'c');e_1_1 = clipboard.paste();e_1_1 = int(e_1_1)
            pg.click(x=852, y=527);pg.hotkey('ctrl', 'c');e_1_h = clipboard.paste();e_1_h = int(e_1_h)
            pg.click(x=899, y=527);pg.hotkey('ctrl', 'c');e_1_m = clipboard.paste();e_1_m = int(e_1_m)

        pg.click(x=892, y=549);pg.hotkey('ctrl', 'c');e_2 = clipboard.paste();e_2 = int(e_2)
        if e_2==1:
            pg.click(x=892, y=575);pg.hotkey('ctrl', 'c');e_2_1 = clipboard.paste();e_2_1 = int(e_2_1)
            pg.click(x=852, y=595);pg.hotkey('ctrl', 'c');e_2_h = clipboard.paste();e_2_h = int(e_2_h)
            pg.click(x=892, y=595);pg.hotkey('ctrl', 'c');e_2_m = clipboard.paste();e_2_m = int(e_2_m)

        E_1=e_1_1*(e_1_h*60+e_1_m)*2 + e_2_1*(e_2_h*60+e_2_m)
        E_2=e_1_1+e_2_1

        if E_1>=300 and E_2>=3 :
            Ea=3
            #print("건강증진신체활동", end=' ')
            swresult = swresult + '건강증진 '
        elif E_1>=150 and E_2>=3 :
            Ea=2
            #print("기본신체활동", end=' ')
            swresult = swresult + '기본 '
        else :
            Ea=1
            #print("신체활동부족", end=' ')
            swresult = swresult + '부족 '

        pg.click(x=892, y=639);pg.hotkey('ctrl', 'c');E_3 = clipboard.paste();E_3 = int(E_3)
        if E_3>=3 :
            Eb=1
            #print("근력운동적절", end=' ')
            swresult = swresult + '적절 '
        else :
            Eb=0
            #print("근력운동부족", end=' ')
            swresult = swresult + '부족 '


        if screen.getpixel((655, 830)) != (0, 0, 0):
            pg.click(x=656, y=831)
            time.sleep(0.01)

        pg.click(x=890,y=854)
        time.sleep(0.01)
        if Ea==3 and Eb==1 : pg.press('1'); time.sleep(0.01); pg.press('1')
        if Ea == 3 and Eb == 0: pg.press('1'); time.sleep(0.01); pg.press('0')
        if Ea == 2 and Eb == 1: pg.press('9');
        if Ea == 2 and Eb == 0: pg.press('8');
        if Ea == 1 and Eb == 1: pg.press('7');
        if Ea == 1 and Eb == 0: pg.press('6');

        pg.click(x=892,y=879)
        if Ea==3:
            if Eb==1:
                pg.press('5')
                #print("5 2 2")
                swresult = swresult + '5 2 2     '
            else :
                pg.press('6')
                #print("6 2 2")
                swresult = swresult + '6 2 2     '

        else :
            pg.press('1')
            #print("1 2 2")
            swresult = swresult + '1 2 2     '

        pg.click(x=892,y=904); pg.press('2')
        pg.click(x=892, y=930);pg.press('2')

        return Ea, Eb

    def exercise_sangwhal(screen_sang_1,Ea,Eb,r_1,r_2,r_3,r_4):
        if screen_sang_1.getpixel((664, 263)) != (0, 0, 0):
            pg.click(664,263)
            time.sleep(0.1)

        pg.click(x=860,y=285)
        if Ea==1 and Eb==0 : pg.press('1')
        elif Ea==1 or Eb==0 : pg.press('2')
        else : pg.press('3');pg.click(x=694,y=438)

        if Ea==1 or Ea==2 : pg.click(x=697,y=332)
        if Eb==0 : pg.click(x=818,y=435)

        pg.click(x=858,y=502); pg.press('2')
        pg.click(x=858, y=569);   pg.press('2')

        pg.click(x=806,y=616)
        if r_1 == 1: pg.click(x=690,y=648)
        if r_2 == 1: pg.click(x=807, y=644)
        if r_3 == 1: pg.click(x=690, y=707)
        if r_4 == 1: pg.click(x=690, y=619)

    def yongyang_chu(screen_chu_2):
        global swresult
        ##yongyang a_1 to a_11, A_1 is yeongyang_jeomsu
        pg.click(x=291, y=279);    pg.hotkey('ctrl', 'c');    a_1 = clipboard.paste()
        pg.click(x=291, y=307);    pg.hotkey('ctrl', 'c');    a_2 = clipboard.paste()
        pg.click(x=291, y=334);    pg.hotkey('ctrl', 'c');    a_3 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 3);    pg.hotkey('ctrl', 'c');    a_4 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 4);    pg.hotkey('ctrl', 'c');    a_5 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 5);    pg.hotkey('ctrl', 'c');    a_6 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 6);    pg.hotkey('ctrl', 'c');    a_7 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 7);    pg.hotkey('ctrl', 'c');    a_8 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 8);    pg.hotkey('ctrl', 'c');    a_9 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 9);    pg.hotkey('ctrl', 'c');    a_10 = clipboard.paste()
        pg.click(x=291, y=279 + 27 * 10);    pg.hotkey('ctrl', 'c');    a_11 = clipboard.paste()

        a_1 = int(a_1);    a_2 = int(a_2);    a_3 = int(a_3);    a_4 = int(a_4);    a_5 = int(a_5)
        a_6 = int(a_6);    a_7 = int(a_7);    a_8 = int(a_8);    a_9 = int(a_9);    a_10 = int(a_10)
        a_11 = int(a_11)

        if screen_chu_2.getpixel((46, 593)) != (0, 0, 0):
            pg.click(46, 593)
            time.sleep(0.2)

        pg.click(x=291, y=617)
        pg.hotkey('ctrl', 'c')
        A_1 = clipboard.paste()
        A_1 = int(A_1)
        if A_1==1:
            #print("영양양호", end=' ')
            swresult = swresult + '양호 '
        elif A_1==2:
            #print("영양보통", end=' ')
            swresult = swresult + '보통 '
        else :
            #print("영양불량", end=' ')
            swresult = swresult + '불량 '

        if a_1 != 1:
            pg.click(x=50, y=667)
            #print("1-1", end=' ')
            swresult = swresult + '1-1 '
        if a_2 != 1:
            pg.click(x=126, y=667)
            #print("1-2", end=' ')
            swresult = swresult + '1-2 '
        if a_3 != 1 or a_4 != 1:
            pg.click(x=219, y=667)
            #print("1-3", end=' ')
            swresult = swresult + '1-3'
        #print('   ')
        swresult = swresult + '   '
        if a_6 != 3:
            pg.click(x=50, y=714)
            #print("2-1", end=' ')
            swresult = swresult + '2-1 '
        if a_7 != 3:
            pg.click(x=126, y=714)
            #print("2-2", end=' ')
            swresult = swresult + '2-2 '
        if a_8 != 3:
            pg.click(x=219, y=714)
            #print("2-3", end=' ')
            swresult = swresult + '2-3 '
        #print('   ')
        swresult = swresult + '   '
        if a_9 != 1:
            pg.click(x=50, y=764)
            #print("3-1", end=' ')
            swresult = swresult + '3-1 '
        if a_10 != 1:
            pg.click(x=219, y=764)
            #print("3-2", end=' ')
            swresult = swresult + '3-2 '
        swresult = swresult + '   '
        print("")

        return A_1, a_1, a_2, a_3, a_5, a_7, a_8, a_9, a_10, a_11

    def yongyang_sangwhal(screen_sang_2,A_1, a_1, a_2, a_3, a_5, a_7, a_8, a_9, a_10, a_11, r_1, r_2, r_3, r_4):

        if screen_sang_2.getpixel((44, 258)) != (0, 0, 0):
            pg.click(44, 258)
            time.sleep(0.1)

        pg.click(x=189,y=286)
        if A_1==1:pg.press('3')
        if A_1==2:pg.press('2')
        if A_1==3:pg.press('1')

        if a_1 != 1: pg.click(x=48, y=344)
        if a_2 != 1: pg.click(x=181, y=344)
        if a_3 != 1: pg.click(x=48, y=382)
        if a_5 != 1: pg.click(x=181,y=382)
        if a_7 != 3: pg.click(x=181, y=418)
        if a_8 != 3: pg.click(x=48, y=456)
        if a_9 != 1: pg.click(x=181, y=456)
        if a_10 != 1: pg.click(x=48, y=491)
        if a_11 != 1: pg.click(x=181, y=491)

        if r_1==1: pg.click(x=48,y=569)
        if r_2 == 1: pg.click(x=181, y=569)
        if r_3 == 1: pg.click(x=181, y=601)
        if r_4 == 1: pg.click(x=181, y=673)

    def obese_chu(screen_chu_2,sex,drinking):
        global swresult
        if screen_chu_2.getpixel((375, 257)) != (0, 0, 0):        pg.click(376,257)
        else : pg.click(x=376,y=257,clicks=2,interval=0.3)
        pg.click(586,311);pg.hotkey('ctrl','c');o_1=clipboard.paste();o_1=float(o_1)
        pg.click(579,282);pg.hotkey('ctrl','c');o_2=clipboard.paste();o_2=float(o_2)
        pg.click(456, 315);    pg.hotkey('ctrl', 'c');    o_3 = clipboard.paste();

        screen_chu_2 = ImageGrab.grab()
        if screen_chu_2.getpixel((375, 471)) != (0, 0, 0):
            pg.click(375, 471)
            time.sleep(0.01)

        pg.click(604,500)
        pg.press('backspace')

        if o_1>=30 :
            Oa=3
            pg.press('3')
            #print("비만", end=' ')
            swresult = swresult + '비만 '
        elif o_1>=25 :
            Oa=2
            pg.press('2')
            #print("과체중", end=' ')
            swresult = swresult + '과체중 '
        else :
            Oa=1
            pg.press('1')
            #print("정상체중", end=' ')
            swresult = swresult + '정상체중 '
        pg.click(393,669)

        if Oa>=2 :
            pg.click(393,551)
            pg.click(393,579)
            if drinking==1:pg.click(393,610)
            pg.click(393,639)


        if sex==1 :
            if o_2>=90 :
                Ob=1
                #print("복부비만", end=' ')
                swresult = swresult + '복부비만 '
            else : Ob=0
        else :
            if o_2>=85 :
                Ob=1
                #print("복부비반", end=' ')
                swresult = swresult + '복부비만 '
            else : Ob=0
        #print("운동처방")
        swresult = swresult + '운동처방'

        return Oa,Ob,o_3

    def obese_sang(screen_sang_2,Oa,Ob,o_3,smoking,drinking,r_1,r_2,r_3,r_4):
        if screen_sang_2.getpixel((345, 259)) != (0, 0, 0):
            pg.click(345,259)
            time.sleep(0.01)

        pg.click(494,287)
        if Oa==1:pg.press('1')
        elif Oa==2:pg.press('2')
        else:pg.press('3')

        pg.click(491,309)
        if Ob==1:pg.press('1')
        else:pg.press('2')

        o_5=0
        pg.click(492,336)
        if Oa==1 and Ob==0 :o_5=1; pg.press('2')
        elif Oa>=2 and Ob==1 : pg.press('6')
        else : pg.press('4')

        o_3_1=o_3[0]
        o_3_2=o_3[1]
        o_3_float=float(o_3)
        o_3_3=0
        if(o_3_float>=100):o_3_3=o_3[2]


        if o_5==1:
            pg.click(436,360); pg.press('0')

            pg.click(379,384)
            pg.press(o_3_1)
            pg.press(o_3_2)

            if (o_3_float >= 100):pg.press(o_3_3)

            pg.click(379,402); pg.press('0')
            pg.click(444,407); pg.press('0')
            pg.click(497,500)

        else:
            pg.click(436, 360);
            pg.press('1'); pg.press('0')

            o_3_gaewol=o_3_float/10
            o_3_gaewol=int(o_3_gaewol)
            o_3_gaewol_str=str(o_3_gaewol)
            o_3_gaewol_str_1=o_3_gaewol_str[0]
            o_3_gaewol_str_2=0
            if(o_3_gaewol>=10):o_3_gaewol_str_2=o_3_gaewol_str[1]

            o_3_mokpwo=o_3_float-o_3_gaewol
            o_3_mokpwo_str=str(o_3_mokpwo)

            o_3_1 = o_3_mokpwo_str[0]
            o_3_2 = o_3_mokpwo_str[1]
            o_3_3 = 0
            if (o_3_mokpwo >= 100): o_3_3 = o_3_mokpwo_str[2]

            pg.click(379,384)
            pg.press(o_3_1)
            pg.press(o_3_2)
            if (o_3_mokpwo >= 100):pg.press(o_3_3)

            pg.click(379, 402)
            pg.press(o_3_gaewol_str_1)
            if(o_3_gaewol>=10):pg.press(o_3_gaewol_str_2)

            pg.click(444, 407)
            pg.press('1')
            time.sleep(0.01)

            pg.click(367, 451)
            pg.click(497,451)
            pg.click(367,478)
            if smoking==1: pg.click(497,478)
            if drinking==1: pg.click(367,504)
            pg.click(497,504)

            if r_1 == 1: pg.click(x=498, y=686)
            if r_2 == 1: pg.click(x=365, y=659)
            if r_3 == 1: pg.click(x=366, y=708)

    def uul():
        pg.click(875,335);pg.hotkey('ctrl', 'c');u_1 = clipboard.paste();u_1 = int(u_1)

        if u_1==1:
            pg.click(x=107, y=201, clicks=2)
            time.sleep(1)

            pg.click(904,689)
            pg.click(147,801)
            pg.click(794,590)
            pg.click(823,806)

        return u_1

    def professor():
        global dateLee
        global dateSong
        global pjdate

        imagedate = ImageGrab.grab((96, 101, 170, 116))

        text = pytesseract.image_to_string(imagedate, lang='eng')
        date = text[0:10]
        date2 = date[5:7] + date[8:10]
        dateint = int(date2)

        for i in range(len(dateLee)):
            if dateint == dateLee[i]:
                prof = '01'
        for i in range(len(dateSong)):
            if dateint == dateSong[i]:
                prof = '02'



        pg.click(740,264)
        pg.write(prof)
        pg.press('enter')
        time.sleep(0.01)


        pg.click(738,287)
        time.sleep(0.01)
        pg.press('backspace', presses=2)
        time.sleep(0.01)
        pg.write(prof)
        pg.press('enter')

        time.sleep(0.01)
        pg.click(781,310)
        time.sleep(0.1)
        pg.press('backspace',presses=8)
        time.sleep(0.01)
        pg.write(pjdate)



    #guiju_jilwhan
    def guijeo():
        r_1=0;r_2=0;r_3=0;r_4=0
        pg.click(x=107, y=201, clicks=2)
        time.sleep(1)
        screen_1 = ImageGrab.grab()
        if screen_1.getpixel((444,860)) == (0, 65, 132) or screen_1.getpixel((823,862)) == (0, 65, 132):
            r_1=1
        if screen_1.getpixel((444,879)) == (0, 65, 132) or screen_1.getpixel((822,881)) == (0, 65, 132):
            r_2=1
        if screen_1.getpixel((443,902)) == (0, 65, 132) or screen_1.getpixel((537,901)) == (0, 65, 132):
            r_3=1
        if screen_1.getpixel((202,863)) == (0, 65, 132) or screen_1.getpixel((679,916)) == (0, 65, 132):
            r_4=1
        #print('HTN, DM, dyslipidemia, obesity : ',r_1, r_2, r_3, r_4)

        return r_1,r_2,r_3,r_4


    def sangwhalstart():
        global swresult
        #get professor, date

        #get sex
        pg.click(417,303, clicks=2, interval=0.1)
        pg.hotkey('ctrl','c')
        sex_get=clipboard.paste()
        sex_get=sex_get[0]
        if sex_get=="남": sex=1
        else: sex=2

        #get guijeo_jealwhan
        r_1,r_2,r_3,r_4=guijeo()
        #r_1:hypertension, r_2:diabetes, r_3:dyslipidemia, r_4:obesity


        #########chuga_gumjin
        pg.click(x=480,y=203, clicks=2)
        #smoking,drinking,excercise
        pg.click(x=252,y=233, clicks=2)
        time.sleep(0.7)

        smoking=0;drinking=0
        screen=ImageGrab.grab()
        time.sleep(0.1)
        if screen.getpixel((293,573)) == (0,0,0):
            smoking=1
            #print("현재흡연자 ")
            swresult = swresult + '현재흡연자 '
            nicotine,cessation=smoking_chu()
        else :
            #print("비흡연자 1 1")
            swresult = swresult + '비흡연자 1 1     '
        if screen.getpixel((353,664)) == (0,0,0):
            drinking=1
            d_1=drinking_chu()
        else :
            #print("적정음주 1")
            swresult = swresult + '적정음주 1     '
        Ea,Eb=exercise_chu(screen)

        #yongyang,obese,uul
        pg.click(x=711,y=229,clicks=2)
        time.sleep(0.7)
        screen_chu_2=ImageGrab.grab()

        yA_1, ya_1, ya_2, ya_3, ya_5, ya_7, ya_8, ya_9, ya_10, ya_11=yongyang_chu(screen_chu_2)

        Oa,Ob,o_3=obese_chu(screen_chu_2,sex,drinking)
        #u_1=uul()




        ###########sangwhal_seobgwan
        pg.click(x=673,y=204,clicks=2)
        #sangwhal_chubang(1)
        pg.click(x=271,y=229,clicks=2)
        time.sleep(1)
        screen_sang_1=ImageGrab.grab()


        if smoking==1:
            smoking_sangwhal(nicotine,cessation,screen_sang_1)
        if drinking==1:
            drinking_sangwhal(d_1,screen_sang_1,r_1,r_2,r_3,r_4)
        exercise_sangwhal(screen_sang_1,Ea,Eb,r_1,r_2,r_3,r_4)


        #sangwhal_chubang(2)
        pg.click(x=715,y=229,clicks=2)
        time.sleep(1)
        screen_sang_2=ImageGrab.grab()
        yongyang_sangwhal(screen_sang_2,yA_1, ya_1, ya_2, ya_3, ya_5, ya_7, ya_8, ya_9, ya_10, ya_11, r_1, r_2, r_3, r_4)
        obese_sang(screen_sang_2,Oa,Ob,o_3,smoking,drinking,r_1,r_2,r_3,r_4)
        professor()

        #저장은 안함함


    #######################################################################################





    class GLOBAL :
        def __init__(self):
            self.raw = []
            self.mulgilcode = 0
            self.kdsq = 0
            self.bone = 0
            self.uul = 0
            self.sangwhal = 0


        def mulgil(self):
            global ERROR
            global savecode
            imagemulgil = ImageGrab.grab((962, 67, 1016, 104))
            text = pytesseract.image_to_string(imagemulgil, lang='eng')
            #print(text)

            if text == "" :
                self.mulgilcode= 99999
            elif text[:5] == "70001" :
                self.mulgilcode = 70001
            elif text[:5] == "21209" and (text[6:11] == "70001" or text[6:11] == "T0001"):
                self.mulgilcode = 70002
            elif text[:5] == "21209" and text[6].isalnum() == False:
                self.mulgilcode = 21209
            elif text[:4] == "S101" and text[5].isalnum() == False:
                self.mulgilcode = 10000
            elif text[:5] == "21055" and text[6].isalnum() == False:
                self.mulgilcode = 10000
            elif text[:6] == "124005" and text[7].isalnum() == False:
                self.mulgilcode = 10000
            else:
                self.mulgilcode = 0


            if self.mulgilcode == 99999 :
                pg.click(84, 82)
                time.sleep(0.01)
                pg.hotkey('ctrl', 'c')
                iljongCode = clipboard.paste()
                if iljongCode == "A" or iljongCode == "B" or iljongCode == "X" or iljongCode == "C":
                    if code == "일검" : pass
                    else :
                        ERROR = 1
                        savecode = "일검판정"
                elif iljongCode == "H" or iljongCode == "K" :
                    if code == "종검" : pass
                    else :
                        ERROR = 1
                        savecode = "종검판정"
                else :
                    ERROR = 1
                    savecode = "코드오류"
            elif self.mulgilcode == 70001 or self.mulgilcode == 70002 or self.mulgilcode == 21209 or self.mulgilcode ==10000:
                if code == "야간" : pass
                else :
                    ERROR = 1
                    savecode = "야간판정"
            elif self.mulgilcode == 0 :
                if code == "물질" : pass
                else :
                    ERROR = 1
                    savecode = "물질판정"


        def getraw(self):
            global ERROR
            global savecode
            pg.click(1053,521)
            pg.click(74,242)
            pg.dragRel(360, 0, 0.3, button='left')
            pg.hotkey('ctrl', 'c')

            rawpd = pd.read_clipboard()
            if rawpd.shape[1] != 5 :
                ERROR = 1
                savecode = "시간부족"

            countnan = 0
            for i in range (len(rawpd)):
                hangmok = rawpd.iloc[i][0]

                if hangmok == " 흉부방사선비고" : pass
                elif hangmok[:5] == " Full" : pass
                #elif hangmok == "요단백" : pass
                elif hangmok == " 요단백" and pd.isnull(rawpd.iloc[i][1]) : savecode = "요단백없음"
                elif hangmok == " 인지기능장애" : self.kdsq = 1
                elif hangmok == " 골밀도의뢰기관기호" : self.bone = 1
                elif hangmok == " 정신건강검사(우울증)" : self.uul = 1
                elif hangmok == " 생활습관평가" : self.sangwhal = 1
                else :
                    if pd.isnull(rawpd.iloc[i][1]) == 1 :
                        savecode = hangmok + '미입력'
                        countnan = countnan + pd.isnull(rawpd.iloc[i][1])

            if countnan and mode==2 > 0:
                ERROR = 1
                savecode = "문진미입력"
                return 0


            self.raw = pd.DataFrame.to_numpy(rawpd)



    class Click :
        def __init__(self):
            pass
        def gitarclick(self, gitarment):
            time.sleep(0.1)
            if self.screen.getpixel((26,862)) == (255,0,0):
                pg.click(26,862)
                time.sleep(0.1)
            pg.click(492,709)
            pg.hotkey('ctrl','home')
            clipboard.copy(gitarment)
            pg.hotkey('ctrl', 'v')
            pg.press('enter')

            pg.click(1031,397)

        def ujilwhanclick(self, ujilwhanment):

            pg.click(711, 562)
            pg.hotkey('ctrl','home')
            clipboard.copy(ujilwhanment)
            pg.hotkey('ctrl', 'v')
            pg.press('enter')

            pg.click(1031,397)



        def hbover(self):
            self.gitarclick(importgitar.iloc[5, 2])
        def weightunder(self):
            #10.기타질환관리, 기타질환관리세부항목 추가
            if self.screen.getpixel((319,937)) != (0, 65, 132):
                pg.click(319,937)
                pg.click(669,806)
                pg.press('2')
                pg.press('enter')
            #정상 A인 경우 지우기

            self.gitarclick(importgitar.iloc[15,2])

        def bpunder(self):
            self.gitarclick(importgitar.iloc[17,2])

        def glucoseunder(self):
            self.gitarclick(importgitar.iloc[16, 2])

        def hepatitisB(self, result):
            if result == 'carrier' :
                #문진등록 확인하기
                pg.click(296,198)
                time.sleep(0.5)

                #B형간염 항원 보유자입니까?
                pg.click(229,633)
                pg.hotkey('ctrl', 'c')
                viral = clipboard.paste()
                if viral == '1' :
                    self.gitarclick(importgitar.iloc[6,2])
                if viral == '2' or '3' :
                    self.gitarclick(importgitar.iloc[7,2])
            if result == 'vaccine' :
                self.gitarclick(importgitar.iloc[8,2])

        def naksang(self):
            self.gitarclick(importgitar.iloc[9,2])

        def uulclick(self):
            self.gitarclick(importgitar.iloc[10, 2])

        def kdsqclick(self):
            self.gitarclick(importgitar.iloc[13, 2])

        def gitarsebuclick(self,gitar4):
            if self.screen.getpixel((319,937)) != (0, 65, 132):
                pg.click(319,937)
            pg.click(669,806)
            pg.press('4')
            pg.press('enter')
            pg.click(516,825)
            clipboard.copy(gitar4)
            pg.hotkey('ctrl', 'v')




    class Panjung(GLOBAL, Click) :
        def __init__(self):
            super().__init__()
            self.screenzero = ImageGrab.grab()
            self.getraw()
            self.indv = 0

        def screencapture(self):
            self.screen = ImageGrab.grab()

        def flask(self):
            if self.screenzero.getpixel((765,353)) == (217, 217, 255):
                pass
            else :
                pg.click(689,384, clicks=2)

            time.sleep(0.05)
            pg.click(789,349) #플라스크 누르기

        def indv_or_grp(self):

            if code == "일검" or code == "종검":
                pg.click(304,90)
                pg.hotkey('ctrl', 'c')
                work = clipboard.paste()
                screenilgum = ImageGrab.grab()
                if work == "성인병" or work == "종검개인" :
                    self.indv = 1
                    if screenilgum.getpixel((509, 361)) == (0, 0, 0):
                        pg.click(507,363)
                        pg.click(507,376)

                    self.flask()
                    time.sleep(7)

                else :

                    if screenilgum.getpixel((509, 361)) != (0, 0, 0) :
                        pg.click(507,363)
                        pg.click(507,376)

                    self.flask()
                    time.sleep(4)

            else:
                self.flask()
                time.sleep(4)

        def ujilwhanadd(self):

            #4. 폐결핵
            if self.screen.getpixel((444, 918)) == (0, 65, 132) :
                self.ujilwhanclick(importujilwhan.iloc[8,2])


            #3.이상지질
            if self.screen.getpixel((443, 902)) == (0, 65, 132) :
                check = 0
                for i in range(len(self.raw)):
                    if self.raw[i, 0] == " 총콜레스테롤":
                        check = 1
                        if int(self.raw[i,1]) < 240 and int(self.raw[i+1,1]) >= 40 and int(self.raw[i+2,1]) <160 and int (self.raw[i+3,1]) <200 :
                            self.ujilwhanclick(importujilwhan.iloc[7,2])
                        else :
                            self.ujilwhanclick(importujilwhan.iloc[6,2])

                if check == 0:
                    self.ujilwhanclick(importujilwhan.iloc[5,2])


            #2.당뇨
            if self.screen.getpixel((444, 879)) == (0, 65, 132):
                for i in range(len(self.raw)):
                    if self.raw[i, 0] == " 식전혈당":
                        if int(self.raw[i, 1]) < 126:
                            self.ujilwhanclick(importujilwhan.iloc[2, 2])
                        else:
                            self.ujilwhanclick(importujilwhan.iloc[3, 2])


            #1.고혈압
            if self.screen.getpixel((444, 860)) == (0, 65, 132) :
                for i in range(len(self.raw)):
                    if self.raw[i, 0] == " 혈압(수축기)":
                        if int(self.raw[i, 1]) < 140 and int(self.raw[i + 1, 1]) < 90:
                            self.ujilwhanclick(importujilwhan.iloc[0,2])
                        else :
                            self.ujilwhanclick(importujilwhan.iloc[1,2])

            if self.indv == 0 :
                if self.screen.getpixel((27,900)) == (255,0,0) and (self.screen.getpixel((27,917)) == (255,0,0) or self.screen.getpixel((25,939)) ==(255,0,0) ) :
                    pg.click(860,495)
                    pg.press('4')
                    pg.press('enter')
                    pg.click(1040,577)


        def gitaradd(self):
            global savecode
            global gitar1
            global gitar2
            global gitar3
            global gitar4
            pg.click(1068,534)
            time.sleep(0.2)

            for i in range (len(self.raw)) :
                if self.raw[i,0] == " 혈색소" :
                    if self.raw[i,4][0] == "남" :
                        sexnum = 1
                    else :
                        sexnum = 2
                    if float(self.raw[i,1]) > 17.5-sexnum:
                        self.hbover()
                        gitar1 = 1

                if self.raw[i,0] == " 체질량지수" :
                    if float(self.raw[i,1]) < 18.5 :
                        self.weightunder()
                        gitar2 = 1

                if self.raw[i,0] == " 혈압(수축기)" :
                    if int(self.raw[i,1]) < 90 or int(self.raw[i+1,1]) < 60 :
                        self.bpunder()
                        gitar4 = gitar4 + " 저혈압"

                if self.raw[i,0] == " 식전혈당" :
                    if int(self.raw[i,1]) < 50 :
                        self.glucoseunder()
                        gitar4 = gitar4 + " 저혈당"

                if self.raw[i,0] == " 간염검사결과" :
                    if self.raw[i,1][0] == "3" : self.hepatitisB('carrier')
                    if self.raw[i,1][0] == "2" : self.hepatitisB('vaccine')
                    if self.raw[i,1][0] == "4" : savecode = savecode + 'b형간염'

                if self.raw[i,0] == " 평형성 (눈감은상태)" :

                    if int(self.raw[i,1]) <= 14 or int(self.raw[i+1,1]) <=19 :
                        self.naksang()
                        gitar4 = gitar4 + " 낙상위험"
            if self.sangwhal == 1 :
                sangwhalstart()
            if self.uul == 1 :
                pg.click(485,200)
                time.sleep(0.5)

                pg.click(875, 335)
                pg.hotkey('ctrl', 'c')
                u_1 = clipboard.paste()
                u_1 = int(u_1)

                pg.click(108,201)
                time.sleep(0.5)

                if u_1 == 1 : self.uulclick()


            if self.kdsq == 1 :
                pg.click(485, 200)
                time.sleep(0.5)

                pg.click(875, 437)
                pg.hotkey('ctrl', 'c')
                k_1 = clipboard.paste()



                pg.click(108, 201)
                time.sleep(0.5)
                if k_1 == '2' :
                    self.kdsqclick()

            if gitar4 != "" :
                if gitar1 + gitar2 + gitar3 == 0 :
                    self.gitarsebuclick(gitar4)







        def d2add(self):
            global savecode
            AA = 0
            JJ = 0
            DD = 0
            II = 0
            if self.screen.getpixel((537,861)) == (0, 65, 132) :
                pg.click(653,862)
                AA = 1
            if self.screen.getpixel((538,879)) == (0, 65, 132) :
                pg.click(653,881)
                JJ = 1
            if self.screen.getpixel((680,859)) == (0, 65, 132) :
                pg.click(798,861)
                DD = 1


            II=0
            '''
            for i in range(len(self.raw)):
                if self.raw[i, 0] == " Chest PA":
                    if pd.isnull(self.raw[i,1][0]) : pass
                    elif self.raw[i,1][0] == "9" :
                        savecode = "엑스레이확인"
                        pg.click(679, 939)
                        pg.click(798, 937)
                        II = 1
            '''

            if JJ+II > 0 :
                pg.click(692,522)
                time.sleep(0.2)
                if JJ == 1:
                    pg.click(92,262)
                if II == 1:
                    pg.click(92,343)


                pg.click(113,723)


    class Save :
        def __init__(self):
            self.date = ""
            self.dateint = 0
        def getdate(self):
            imagedate = ImageGrab.grab((96, 101, 170, 116))

            text = pytesseract.image_to_string(imagedate, lang='eng')
            self.date = text[0: 10]
            date2 = self.date[5:7] + self.date[8:10]
            self.dateint = int(date2)

        def getscreenresult(self):
            global gitar1
            global gitar2
            global gitar3
            global gitar4

            self.result1 = ''
            self.result2 = ''
            self.result3 = ''
            self.result4 = ''
            self.result5 = ''
            scr = ImageGrab.grab()
            if scr.getpixel((27, 861)) == (255, 0, 0) : self.result1 = '정상'
            ###########################
            if scr.getpixel((203, 861)) == (0, 65, 132): self.result2 = self.result2 + '1 '
            if scr.getpixel((203, 879)) == (0, 65, 132): self.result2 = self.result2 + '2 '
            if scr.getpixel((202, 900)) == (0, 65, 132): self.result2 = self.result2 + '3 '
            if scr.getpixel((203, 920)) == (0, 65, 132): self.result2 = self.result2 + '4 '
            if scr.getpixel((202, 937)) == (0, 65, 132): self.result2 = self.result2 + '5 '

            if scr.getpixel((319, 862)) == (0, 65, 132): self.result2 = self.result2 + '6 '
            if scr.getpixel((318, 881)) == (0, 65, 132): self.result2 = self.result2 + '7 '
            if scr.getpixel((320, 898)) == (0, 65, 132): self.result2 = self.result2 + '8 '
            if scr.getpixel((317, 917)) == (0, 65, 132): self.result2 = self.result2 + '9 '

            ###혈색소, 저체중, 시력저하, '(1)' '(2)' '(3)' '(4)'###
            if gitar1 == 1: self.result2 = self.result2 + '(1) '
            if gitar2 == 1: self.result2 = self.result2 + '(2) '
            if gitar3 == 1: self.result2 = self.result2 + '(3) '
            if gitar4 != "": self.result2 = self.result2 + '(4) '
            ############################

            if scr.getpixel((538, 861)) == (0, 65, 132): self.result3 = self.result3 + '1 '
            if scr.getpixel((538, 878)) == (0, 65, 132): self.result3 = self.result3 + '2 '
            if scr.getpixel((538, 898)) == (0, 65, 132): self.result3 = self.result3 + '3 '
            if scr.getpixel((537, 918)) == (0, 65, 132): self.result3 = self.result3 + '4 '
            if scr.getpixel((538, 936)) == (0, 65, 132): self.result3 = self.result3 + '5 '

            if scr.getpixel((679, 859)) == (0, 65, 132): self.result3 = self.result3 + '6 '
            if scr.getpixel((679, 882)) == (0, 65, 132): self.result3 = self.result3 + '7 '
            if scr.getpixel((678, 901)) == (0, 65, 132): self.result3 = self.result3 + '8 '
            if scr.getpixel((679, 919)) == (0, 65, 132): self.result3 = self.result3 + '9 '
            if scr.getpixel((678, 938)) == (0, 65, 132): self.result3 = self.result3 + '10 '
            ##########################
            if scr.getpixel((823, 860)) == (0, 65, 132): self.result4 = self.result4 + '1 '
            if scr.getpixel((823, 879)) == (0, 65, 132): self.result4 = self.result4 + '2 '
            #############################
            if scr.getpixel((444, 860)) == (0, 65, 132): self.result5 = self.result5 + '1 '
            if scr.getpixel((444, 879)) == (0, 65, 132): self.result5 = self.result5 + '2 '
            if scr.getpixel((443, 902)) == (0, 65, 132): self.result5 = self.result5 + '3 '
            if scr.getpixel((444, 918)) == (0, 65, 132): self.result5 = self.result5 + '4 '
            ##############################


        def getlist(self):
            global ERROR
            global savecode
            global exportdf
            global swresult

            pg.click(1286, 632)
            pg.click(299,89);pg.hotkey('ctrl','c');work=clipboard.paste();time.sleep(0.05)
            pg.click(463,89);pg.hotkey('ctrl','c');id=clipboard.paste();time.sleep(0.05)
            pg.click(558,89);pg.hotkey('ctrl','c');name=clipboard.paste()

            if savecode == "" and ERROR == 0:
                savecode = "완료"
            savelist = [[work, self.date, name, savecode, self.result1, self.result2, self.result3, self.result4, self.result5,swresult]]
            print(savelist)

            appdf = pd.DataFrame(data=savelist, columns=['사업장','검진일','성함','판정여부','정상A','정상B','일반질환','만성의심','유질환자','생활습관'])
            exportdf = exportdf.append(appdf)
        def saveclick(self):
            global dateLee
            global dateSong
            global pjdate

            self.getdate()
            self.getscreenresult()
            self.getlist()
            if datemode==1 :
                pg.click(557,438)
                pg.press('delete',presses = 8)
                pg.write(pjdate)


            if profmode==1 :
                for i in range (len(dateLee)) :
                    if self.dateint == dateLee[i] :
                        prof = '01'
                for i in range (len(dateSong)) :
                    if self.dateint == dateSong[i] :
                        prof = '02'

                pg.click(781,437)
                pg.write(prof)
                pg.press('enter')

            pg.click(1647,243)
            time.sleep(2)

            screenerror = ImageGrab.grab()
            if screenerror.getpixel((838,533)) == (66,106,208) :
                pg.click(1066,596) #일반/생애 1차 건강검진 대상자가 아닙니다
                #pg.click(1647,243)






    print("로딩완료")
    if mode == 1 :
        while True :
            key = keyboard.read_hotkey(suppress=False)

            if key == 'f9' :

                ERROR = 0
                savecode = ""
                gitar1 = 0
                gitar2 = 0
                gitar3 = 0
                gitar4 = ""
                swresult = ''

                test = GLOBAL()
                #test.mulgil()
                if ERROR == 1:
                    pass
                else:
                    pjtest = Panjung()  # test.getraw 불러옴
                    if ERROR == 1:
                        pass
                    else:
                        pjtest.indv_or_grp()  # pjtest.flask() 불러옴
                        pg.click(1042, 551)
                        pjtest.screencapture()
                        pjtest.gitaradd()
                        pjtest.ujilwhanadd()
                        pjtest.d2add()

                    del pjtest
                save = Save()
                save.saveclick()


                del test

                pg.click(1066,596) #빈 곳 한 번 클릭

            elif key == 'alt+d' :
                print(exportdf)
                exportdf.to_excel('C:\\Users\\DELL3040\\Desktop\\ilgumdf.xlsx')
            else :
                time.sleep(0.1)

    elif mode == 2 :
        print("강제중지하려면 esc 계속 누르고 있기 - 한사람 끝나면 강제중지됨")
        for i in range(rep) :
            time.sleep(1)
            print(i+1, end='')
            ERROR = 0
            savecode = ""
            gitar1 = 0
            gitar2 = 0
            gitar3 = 0
            gitar4 = ""
            swresult = ""

            e = ""

            try :
                if keyboard.is_pressed('esc') :
                    print('강제종료됨')
                    break

                test = GLOBAL()
                test.mulgil()
                if ERROR == 1:
                    pass
                else:
                    pjtest = Panjung()  # test.getraw 불러옴
                    if ERROR == 1:
                        pass
                    else:
                        pjtest.indv_or_grp()  # pjtest.flask() 불러옴
                        pg.click(1042, 551)
                        pjtest.screencapture()
                        pjtest.gitaradd()
                        pjtest.ujilwhanadd()
                        pjtest.d2add()

            except Exception as e :
                del pjtest

                ERROR = 1
                print(e)
                savecode = e
            finally :
                pg.click(103,199) #1차결과등록으로
                time.sleep(0.5)

                save = Save()
                save.saveclick()
                time.sleep(1)

                del test

                pg.click(1107, 591)  # 빈 곳 한 번 클릭
                time.sleep(1)

        try :
            exportdf.to_excel('C:\\Users\\DELL3040\\Desktop\\ilgumdf.xlsx')
        except IOError :
            print('ilgumdf.xlsx 파일을 닫으세요')
        print('프로그램종료')


    else :
        print("change mode")

