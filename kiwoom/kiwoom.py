import sys
import os
from PyQt5.QAxContainer import * # C++ api를 가져올 수 있는 container
from PyQt5.QtCore import *
from config.errCode import *
from beautifultable import BeautifulTable

class Kiwoom(QAxWidget): #QAxWidget 상속
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop() # 로그인 담당 이벤트 루프
        self.get_deposit_loop = QEventLoop() # 예수금 담당 이벤트 루프
        self.get_account_evaluation_balance_loop = QEventLoop() # 계좌 잔고 조회 이벤트 루프


        #계좌 관련 함수
        self.account_number = None
        self.total_buy_money = 0
        self.total_evaluation_money = 0
        self.total_evaluation_profit_and_loss_money = 0
        self.total_yield = 0
        self.account_stock_dict = {}

        # 예수금 관련 변수
        self.deposit = 0
        self.withdraw_deposit = 0
        self.order_deposit = 0

        # 화면 번호
        self.screen_my_account = "1000"

        #초기작업
        self.create_kiwoom_instance()
        self.event_collection()  # 이벤트와 슬롯을 메모리에 먼저 생성.
        self.login()
        input() #아무키나 눌러 다음 수행
        self.get_account_info() # 계좌 번호만 얻어오기
        self.get_deposit_info()  # 예수금 관련된 정보 얻어오기
        self.get_account_evaluation_balance()  # 계좌평가잔고내역 얻어오기

        self.menu()

    # COM 오브젝트 생성.
    def create_kiwoom_instance(self):
        # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
        #컴퓨터\HKEY_CLASSES_ROOT\KHOPENAPI.KHOpenAPICtrl.1\CLSID
        #Control 식별자(A1574A0D-6BFA-4BD7-9020-DED88711818D)
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_collection(self):
        #   [OnEventConnect()이벤트]
          
        #   OnEventConnect(
        #   long nErrCode   // 로그인 상태를 전달하는데 자세한 내용은 아래 상세내용 참고
        #   )
          
        #   로그인 처리 이벤트입니다. 성공이면 인자값 nErrCode가 0이며 에러는 다음과 같은 값이 전달됩니다.
 
        #   nErrCode별 상세내용
        #   -100 사용자 정보교환 실패
        #   -101 서버접속 실패
        #   -102 버전처리 실패
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        #   [OnReceiveTrData() 이벤트]
          
        #   void OnReceiveTrData(
        #   BSTR sScrNo,       // 화면번호
        #   BSTR sRQName,      // 사용자 구분명
        #   BSTR sTrCode,      // TR이름
        #   BSTR sRecordName,  // 레코드 이름
        #   BSTR sPrevNext,    // 연속조회 유무를 판단하는 값 0: 연속(추가조회)데이터 없음, 2:연속(추가조회) 데이터 있음
        #   LONG nDataLength,  // 사용안함.
        #   BSTR sErrorCode,   // 사용안함.
        #   BSTR sMessage,     // 사용안함.
        #   BSTR sSplmMsg     // 사용안함.
        #   )
          
        #   요청했던 조회데이터를 수신했을때 발생됩니다.
        #   수신된 데이터는 이 이벤트내부에서 GetCommData()함수를 이용해서 얻어올 수 있습니다.
        self.OnReceiveTrData.connect(self.tr_slot)  # 트랜잭션 요청 관련 이벤트
        
    def login(self):
        #   수동 로그인설정인 경우 로그인창을 출력.
        #   자동로그인 설정인 경우 로그인창에서 자동으로 로그인을 시도합니다.
        self.dynamicCall("CommConnect()")  # 시그널 함수 호출.
        self.login_event_loop.exec_()

    def login_slot(self, err_code):
        if err_code == 0:
            print("로그인에 성공하였습니다.")
        else:
            # KOA Studio나 키움증권 개발가이드에 따르면, 로그아웃을 수행하는 함수는 더이상 지원하지 않으므로 강제종료 로직 필요. 
            os.system('cls')
            print("로그인에 실패하였습니다.")
            print("에러 내용 :", errors(err_code)[1])
            sys.exit(0) # 프로그램 종료
        self.login_event_loop.exit()

    def get_account_info(self):
        #   로그인 후 사용할 수 있으며 인자값에 대응하는 정보를 얻을 수 있습니다.
        #   "ACCOUNT_CNT" : 보유계좌 갯수를 반환합니다.
        #   "ACCLIST" 또는 "ACCNO" : 구분자 ';'로 연결된 보유계좌 목록을 반환합니다.
        #   "USER_ID" : 사용자 ID를 반환합니다.
        #   "USER_NAME" : 사용자 이름을 반환합니다.
        #   "GetServerGubun" : 접속서버 구분을 반환합니다.(1 : 모의투자, 나머지 : 실거래서버)
        #   "KEY_BSECGB" : 키보드 보안 해지여부를 반환합니다.(0 : 정상, 1 : 해지)
        #   "FIREW_SECGB" : 방화벽 설정여부를 반환합니다.(0 : 미설정, 1 : 설정, 2 : 해지)
        account_list = self.dynamicCall("GetLoginInfo(QString)","ACCNO")
        account_number = account_list.split(';')[0] #ToDo: 편의상 첫번째 계좌로 선택
        self.account_number = account_number
        print("Debug get_account_info(): ","계좌번호: ",self.account_number)
    
    def menu(self):
        sel = ""
        while True:
            os.system('cls') # 화면을 깨끗하게 지워줌
            print("1. 현재 로그인 상태 확인")
            print("2. 개인 정보 조회")
            print("3. 예수금 조회")
            print("4. 계좌 잔고 조회")
            print("Q. 프로그램 종료")
            sel = input("=> ")

            if sel == "Q" or sel == "q":
                sys.exit(0) # 프로그램 종료
            
            if sel == "1":
                self.print_login_connect_state()
            elif sel == "2":
                self.print_my_info()
            elif sel == '3':
                self.print_get_deposit_info()
            elif sel == "4":
                self.print_get_account_evaulation_balance_info()

    def print_login_connect_state(self):
        os.system('cls')
        #   서버와 현재 접속 상태를 알려줍니다.
        #   리턴값 1:연결, 0:연결안됨
        isLogin = self.dynamicCall("GetConnectState()")
        if isLogin == 1:
            print("\n현재 계정은 로그인 상태입니다.\n")
        else:
            print("\n현재 계정은 로그아웃 상태입니다.\n")
        input() #아무키나 눌러 다음 수행

    def print_my_info(self):
        os.system('cls') # 화면을 깨끗하게 지워줌
        #   로그인 후 사용할 수 있으며 인자값에 대응하는 정보를 얻을 수 있습니다.
        #   "ACCOUNT_CNT" : 보유계좌 갯수를 반환합니다.
        #   "ACCLIST" 또는 "ACCNO" : 구분자 ';'로 연결된 보유계좌 목록을 반환합니다.
        #   "USER_ID" : 사용자 ID를 반환합니다.
        #   "USER_NAME" : 사용자 이름을 반환합니다.
        #   "GetServerGubun" : 접속서버 구분을 반환합니다.(1 : 모의투자, 나머지 : 실거래서버)
        #   "KEY_BSECGB" : 키보드 보안 해지여부를 반환합니다.(0 : 정상, 1 : 해지)
        #   "FIREW_SECGB" : 방화벽 설정여부를 반환합니다.(0 : 미설정, 1 : 설정, 2 : 해지)
        user_name = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        user_id = self.dynamicCall("GetLoginInfo(QString)", "USER_ID")
        account_count = self.dynamicCall("GetLoginInfo(QString)", "ACCOUNT_CNT")

        print(f"\n이름 : {user_name}")
        print(f"\nID : {user_id}")
        print(f"보유 계좌 수 : {account_count}")
        print(f"계좌번호 : {self.account_number}\n") # TODO: 나의 계좌번호:  8017996111
        input() #아무키나 눌러 다음 수행

    def print_get_deposit_info(self):
        os.system('cls')
        print(f"\n예수금 : {self.deposit}원")
        print(f"출금 가능 금액 : {self.withdraw_deposit}원")
        print(f"주문 가능 금액 : {self.order_deposit}원")
        input() #아무키나 눌러 다음 수행

    def print_get_account_evaulation_balance_info(self):
        os.system('cls')
        print("\n<싱글데이터>")
        print(f"총매입금액 : {self.total_buy_money}원")
        print(f"총평가금액 : {self.total_evaluation_money}원")
        print(f"총평가손익금액 : {self.total_evaluation_profit_and_loss_money}원")
        print(f"총수익률 : {self.total_yield}%")

        table = self.make_table()
        print("<멀티 데이터>")
        print(f"보유 종목 수 : {len(self.account_stock_dict)}개")
        print(table)
        input() #아무키나 눌러 다음 수행


    def make_table(self):
        table = BeautifulTable()
        table = BeautifulTable(maxwidth=150)
        for stock_code in self.account_stock_dict:
            stock = self.account_stock_dict[stock_code]
            stockList = []
            for key in stock:
                output = None

                if key == "종목명":
                    output = stock[key]
                elif key == "수익률(%)":
                    output = str(stock[key]) + "%"
                elif key == "보유수량" or key == "매매가능수량":
                    output = str(stock[key]) + "개"
                else:
                    output = str(stock[key]) + "원"
                stockList.append(output)
            table.rows.append(stockList)
        table.columns.header = ["종목명", "평가손익",
                                "수익률", "매입가", "보유수량", "매매가능수량", "현재가"]
        table.rows.sort('종목명')
        return table
    #   [SetInputValue() 함수]
        
    #   SetInputValue(
    #   BSTR sID,     // TR에 명시된 Input이름
    #   BSTR sValue   // Input이름으로 지정한 값
    #   )
        
    #   조회요청시 TR의 Input값을 지정하는 함수입니다.
    #   CommRqData 호출 전에 입력값들을 셋팅합니다.
    #   각 TR마다 Input 항목이 다릅니다. 순서에 맞게 Input 값들을 셋팅해야 합니다.
        
    #   ------------------------------------------------------------------------------------------------------------------------------------
        
    #   예)
    #   [OPT10081 : 주식일봉차트조회요청예시]
        
    #   SetInputValue("종목코드"	,  "039490"); // 첫번째 입력값 설정
    #   SetInputValue("기준일자"	,  "20160101");// 두번째 입력값 설정
    #   SetInputValue("수정주가구분"	,  "1"); // 세번째 입력값 설정
    #   LONG lRet = CommRqData( "RQName","OPT10081", "0","0600");// 조회요청
        
    #   ------------------------------------------------------------------------------------------------------------------------------------
    #   [CommRqData() 함수]
        
    #   CommRqData(
    #   BSTR sRQName,    // 사용자 구분명 (임의로 지정, 한글지원)
    #   BSTR sTrCode,    // 조회하려는 TR이름
    #   long nPrevNext,  // 연속조회여부
    #   BSTR sScreenNo  // 화면번호 (4자리 숫자 임의로 지정)
    #   )
        
    #   조회요청 함수입니다.
    #   리턴값 0이면 조회요청 정상 나머지는 에러
        
    #   예)
    #   -200 시세과부하
    #   -201 조회전문작성 에러

    def get_deposit_info(self, nPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)",
                         "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000") #TODO 시작프로그램에서 비밀번호를 저장해야함. 
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        #   [CommRqData() 함수]
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "예수금상세현황요청", "opw00001", nPrevNext, self.screen_my_account)

        self.get_deposit_loop.exec_()

    def get_account_evaluation_balance(self, nPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)",
                         "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000") #TODO 시작프로그램에서 비밀번호를 저장해야함. 
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        #   [CommRqData() 함수]
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "계좌평가잔고내역요청", "opw00018", nPrevNext, self.screen_my_account)

        if not self.get_account_evaluation_balance_loop.isRunning():
            self.get_account_evaluation_balance_loop.exec_()

    # sScrNo는 "1000", 
 
    # sRecordName과 sPrevNext는 빈 값으로 날아옵니다. 
    # 여기서 sRecordName은 인자로 넘긴 것이 없으니 빈 값인건 그렇다치더라도 
    # sPrevNext는 "0"이 아니라 빈 값인 ""으로 날아온 것은 조금 의아합니다. 
    # 이것은 현재 조사 중에 있으나, 원인을 알고 계시는 분은 댓글로 남겨 주시면 감사하겠습니다.

    # 화면번호는 서버에 조회나 주문등 필요한 기능을 요청할때 이를 구별하기 위한 키값으로 이해하시면 됩니다. 
    # 0000(혹은 0)을 제외한 임의의 숫자를 사용하시면 되는데 개수가 200개로 한정되어 있기 때문에 
    # 이 개수를 넘지 않도록 관리하셔야 합니다. 
    # 만약 사용하는 화면번호가 200개를 넘는 경우 
    # 조회 결과나 주문 결과에 다른 데이터가 섞이거나 원하지 않는 결과가 나타날 수 있습니다.

    # sPrevNext는 "2"일 경우 다음 페이지가 있다는 의미
    def tr_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        #  [ opw00001 : 예수금상세현황요청 ]

        #  1. Open API 조회 함수 입력값을 설정합니다.
        # 	계좌번호 = 전문 조회할 보유계좌번호
        # 	SetInputValue("계좌번호"	,  "입력값 1");

        # 	비밀번호 = 사용안함(공백)
        # 	SetInputValue("비밀번호"	,  "입력값 2");

        # 	비밀번호입력매체구분 = 00
        # 	SetInputValue("비밀번호입력매체구분"	,  "입력값 3");

        # 	조회구분 = 3:추정조회, 2:일반조회
        # 	SetInputValue("조회구분"	,  "입력값 4");


        #  2. Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # 	CommRqData( "RQName"	,  "opw00001"	,  "0"	,  "화면번호"); 
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            withdraw_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_deposit = int(withdraw_deposit)

            order_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.order_deposit = int(order_deposit)
            self.cancel_screen_number(self.screen_my_account)
            self.get_deposit_loop.exit()


        # [ opw00018 : 계좌평가잔고내역요청 ]

        #  [ 주의 ] 
        #  "수익률%" 데이터는 모의투자에서는 소숫점표현, 실거래서버에서는 소숫점으로 변환 필요 합니다.

        #  1. Open API 조회 함수 입력값을 설정합니다.
        # 	계좌번호 = 전문 조회할 보유계좌번호
        # 	SetInputValue("계좌번호"	,  "8017996111");

        # 	비밀번호 = 사용안함(공백)
        # 	SetInputValue("비밀번호"	,  "0000");

        # 	비밀번호입력매체구분 = 00
        # 	SetInputValue("비밀번호입력매체구분"	,  "00");

        # 	조회구분 = 1:합산, 2:개별
        # 	SetInputValue("조회구분"	,  "2");

        #  2. Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # 	CommRqData( "RQName"	,  "opw00018"	,  "0"	,  "화면번호"); 
        elif sRQName == "계좌평가잔고내역요청":
            if (self.total_buy_money == None or self.total_evaluation_money == None
                or self.total_evaluation_profit_and_loss_money == None or self.total_yield == None):
                total_buy_money = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
                self.total_sell_money = int(total_buy_money)

                total_evaluation_money = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가금액")
                self.total_evaluation_money = int(total_evaluation_money)

                total_evaluation_profit_and_loss_money = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
                self.total_evaluation_profit_and_loss_money = int(
                    total_evaluation_profit_and_loss_money)

                total_yield = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익률(%)")
                self.total_yield = float(total_yield) # 총수익률(%)은 반드시 int가 아닌 float로 형변환

            # 종목의 개수를 GetRepeatCnt() 함수를 통해 최대 20개까지 얻어오고, 
            # cnt번만큼 반복문을 돌면서 종목명, 평가손익, 수익률(%), 매입가, 보유수량, 매매가능수량, 현재가에 관한 데이터를 
            # GetCommData() 함수를 이용하여 얻어 옵니다. 
            # 그리고 각 데이터를 int또는 float로 형변환하여 딕셔너리에 저장하는 것입니다.
            cnt = self.dynamicCall(
                "GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(cnt):
                stock_code = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                stock_code = stock_code.strip()[1:]

                stock_name = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_name = stock_name.strip()  # 필요 없는 공백 제거.

                stock_evaluation_profit_and_loss = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "평가손익")
                stock_evaluation_profit_and_loss = int(
                    stock_evaluation_profit_and_loss)

                stock_yield = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                stock_yield = float(stock_yield)

                stock_buy_money = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                stock_buy_money = int(stock_buy_money)

                stock_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                stock_quantity = int(stock_quantity)

                stock_trade_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                stock_trade_quantity = int(stock_trade_quantity)

                stock_present_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                stock_present_price = int(stock_present_price)

                if not stock_code in self.account_stock_dict:
                    self.account_stock_dict[stock_code] = {}

                self.account_stock_dict[stock_code].update({'종목명': stock_name})
                self.account_stock_dict[stock_code].update(
                    {'평가손익': stock_evaluation_profit_and_loss})
                self.account_stock_dict[stock_code].update(
                    {'수익률(%)': stock_yield})
                self.account_stock_dict[stock_code].update(
                    {'매입가': stock_buy_money})
                self.account_stock_dict[stock_code].update(
                    {'보유수량': stock_quantity})
                self.account_stock_dict[stock_code].update(
                    {'매매가능수량': stock_trade_quantity})
                self.account_stock_dict[stock_code].update(
                    {'현재가': stock_present_price})

            if sPrevNext == "2":
                self.get_account_evaluation_balance("2")
            else:
                self.cancel_screen_number(self.screen_my_account)
                self.get_account_evaluation_balance_loop.exit()

    #   [DisconnectRealData() 함수]
        
    #   DisconnectRealData(
    #   BSTR sScnNo // 화면번호 
    #   )
        
    #   시세데이터를 요청할때 사용된 화면번호를 이용하여 
    #   해당 화면번호로 등록되어져 있는 종목의 실시간시세를 서버에 등록해지 요청합니다.
    #   이후 해당 종목의 실시간시세는 수신되지 않습니다.
    #   단, 해당 종목이 또다른 화면번호로 실시간 등록되어 있는 경우 해당종목에대한 실시간시세 데이터는 계속 수신됩니다.
    def cancel_screen_number(self, sScrNo):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)