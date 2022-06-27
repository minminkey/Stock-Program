import sys
import threading

from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.errorCode import *
from config.kiwoomType import *
from socket import *
from config.log_class import *


buysCode = []


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.realType = RealType()
        self.logging = Logging()

        self.sellCount = 0
        self.account_stock_dict = {}
        # print("Kiwoom start")
        self.logging.logger.debug("Kiwoom() class start")

        ##event loop 실행 변수 모음
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()
        self.sell_order_event_loop = QEventLoop()

        ##매도 수익률 변수
        self.lowRate = 0.98
        self.highRate = 1.02
        self.downRate = 0.995
        # self.downRate = 1.0;
        self.maxTotalPrice = 10000000

        ##계좌 관련 변수

        self.account_num = None         ##계좌번호
        self.deposit = 0                ##예수금
        self.oper_money = 0             ##예수금 중 운영하는 금액
        self.use_money = 0              ##실제 투자에 사용할 금액
        self.oper_money_percent = 1     ##예수금 중 사용할 비율
        self.use_money_percent = 0.1    ##oper_money 중 매수 시 사용할 비율
        self.output_deposit = 0         ##출력가능 금액
        # self.total_buy_money = 0
        self.total_profit_loss_money = 0 ##총평가손익금액
        self.total_profit_loss_rate = 0.0 ##총수익률(%)

        ##요청 스크린 번호
        self.cnt = 0
        self.screen_my_info = "2000"    #계좌 관련 스크린 번호
        self.screen_calculation_stock = "4000"
        self.screen_real_stock = "5000"
        self.screen_meme_stock = "6000"
        self.screen_start_stop_real = "1000"

        ##초기 Setting 함수
        self.get_ocx_instance()         #OCS 방식을 파이썬에 사용할 수 있게 변환해주는 함수
        self.event_slots()              #키움과 연결하기 위한 시그널
        self.real_event_slot()
        self.signal_login_commConnect() #로그인 요청 시그널
        self.get_account_info()         #계좌번호 가져오기
        self.detail_account_info()      #예수금 요청 시그널
        self.detail_account_mystock()
        self.screen_number_setting()
        # self.detail_account_mystock()
        # self.sendOrder("000270")

        QTest.qWait(1000)
        #실시간 수신 관련 함수
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", self.screen_start_stop_real, ' ',
                         self.realType.REALTYPE['장시작시간']['장운영구분'], "0")
        for code in self.account_stock_dict.keys():
            screen_num = self.account_stock_dict[code]["스크린번호"]
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, "1")

        thr = threading.Thread(target=self.socketCommunication)
        thr.start()
        idx = 1
        while True:
            if self.sellCount > 0:
                self.detail_account_info()
                self.sellCount = 0
            if len(buysCode) > 0:
                self.get_stock_value(buysCode[0])
                buysCode.pop(0)
            QTest.qWait(500)

    def socketCommunication(self):
        host = "127.0.0.1"
        port = 12345  # 임의번호

        serverSocket = socket(AF_INET, SOCK_STREAM)  # 소켓 생성
        serverSocket.bind((host, port))  # 생성한 소켓에 설정한 HOST와 PORT 맵핑
        while 1:
            serverSocket.listen(1)  # 맵핑된 소켓을 연결 요청 대기 상태로 전환
            print("대기중입니다")

            connectionSocket, addr = serverSocket.accept()  # 실제 소켓 연결 시 반환되는 실제 통신용 연결된 소켓과 연결주소 할당

            print(str(addr), "에서 접속되었습니다.")  # 연결 완료했다고 알림

            data = connectionSocket.recv(1024)  # 데이터 수신, 최대 1024
            print("매수 코드 :", data.decode("utf-8"))  # 받은 데이터 UTF-8
            buysCode.append(data.decode("utf-8"))
            QTest.qWait(500)
        print("socket 종료")
        serverSocket.close()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
        self.OnReceiveMsg.connect(self.msg_slot)

    def real_event_slot(self):
        self.OnReceiveRealData.connect(self.realdata_slot)

    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        self.logging.logger.debug("스크린: %s, 요청이름: %s, tr코드: %s --- %s" %(sScrNo, sRQName, sTrCode, msg))

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop.exec_()


    def login_slot(self, err_code):
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString", "ACCNO")
        account_num = account_list.split(';')[0]
        self.account_num = account_num


    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString", "예수금상세현황요청", "opw00001", sPrevNext, self.screen_my_info)
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)
        self.detail_account_info_event_loop.exec_()

    #Trading Data Slot
    def trdata_slot(self, sSrcNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)
            oper_money = float(self.deposit) * self.oper_money_percent
            self.oper_money = int(oper_money)
            self.use_money = int(self.oper_money/self.use_money_percent)
            if self.use_money>self.maxTotalPrice:
                self.use_money = self.maxTotalPrice
            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.output_deposit = int(output_deposit)
            self.stop_screen_cancel(self.screen_my_info)
            self.detail_account_info_event_loop.exit()
        elif sRQName == "주식기본정보요청":
            sCode = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            sCode = str(sCode)
            sCode = sCode[len(sCode)-6:]
            stockValue = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "현재가")
            stockValue = abs(int(stockValue))
            quantity = int((self.use_money) / stockValue)
            print("%s %d원 매수" % (sCode, stockValue*quantity))
            self.sendBuyOrder(sCode, stockValue, quantity)
            self.sell_order_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            self.total_buy_money = int(total_buy_money)
            total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익률(%)")
            self.total_profit_loss_rate = float(total_profit_loss_rate)
            print("총매입금액 : %s 총평가손익금액 : %s 총수익률 : %s" % (self.total_buy_money, self.total_profit_loss_money, self.total_profit_loss_rate))
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price)
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())
                print("종목번호: %s - 종목명: %s - 보유수량: %s - 매입가: %s - 수익률: %s - 현재가: %s" % (code, code_nm, stock_quantity, buy_price, learn_rate, current_price))
                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수일률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"최고가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})
            print("계좌에 가지고 있는 종목은 %s" % rows)
            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

    #실시간 주가 확인
    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            value = self.dynamicCall("GetCommRealData(QString, int)", sCode, fid)
            if value == '0':
                print("장 시작 전")
            elif value == '3':
                print("장 시작")
            elif value == '2':
                print("장 종료, 동시호가로 넘어감")
            elif value == '4':
                print("3:30 장 종료")
                for code in self.account_stock_dict.keys():
                    self.dynamicCall("SetRealRemove(QString, QString)", self.account_stock_dict[code]["스크린번호"], code)
                self.logging.logger.debug("시스템 종료")
                sys.exit()

        elif sRealType == "주식체결":
            value = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['현재가'])
            value = abs(int(value))
            if sCode in self.account_stock_dict:
                self.account_stock_dict[sCode].update({"현재가": value})
                if self.account_stock_dict[sCode].get("최고가") == None:
                    self.account_stock_dict[sCode].update({"최고가": value})
                elif value > self.account_stock_dict[sCode]["최고가"]:
                    self.account_stock_dict[sCode].update({"최고가": value})
                buy = self.account_stock_dict[sCode]["매입가"]
                high = self.account_stock_dict[sCode]["최고가"]
                # print("Real Data 코드 : ", sCode,"현재가 : ", value)
                if (value / buy) < self.lowRate or ((high / buy) > self.highRate and (value / high) < self.downRate):
                    print("조건 만족")
                    self.sendSellOrder(sCode)


    #실시간 확인 중단
    def stop_screen_cancel(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)

    #주가 확인
    def get_stock_value(self, sCode, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", sCode)
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식기본정보요청", "opt10001", "0", 10)
        self.sell_order_event_loop.exec_()

    #매수 주문 전송
    def sendBuyOrder(self, sCode, buy_price, quantity):
        self.logging.logger.debug("종목코드: %s, 매수 단가: %d, 수량: %d, 총: %d" % (sCode, buy_price, quantity, buy_price * quantity))
        if sCode in self.account_stock_dict:
            pass
        else:
            self.account_stock_dict[sCode] = {}
        self.account_stock_dict[sCode].update({"보유수량": quantity})
        self.account_stock_dict[sCode].update({"매입가": buy_price})
        self.account_stock_dict[sCode].update({"최고가": buy_price})
        temp_screen = int(self.screen_real_stock)
        meme_screen = int(self.screen_meme_stock)
        if (self.cnt % 50) == 0:
            temp_screen += 1
            self.screen_real_stock = str(temp_screen)
        if (self.cnt % 50) == 0:
            meme_screen += 1
            self.screen_meme_stock = str(meme_screen)
        self.account_stock_dict[sCode].update({"스크린번호": str(self.screen_real_stock)})
        self.account_stock_dict[sCode].update({"주문용스크린번호": str(self.screen_meme_stock)})
        self.cnt += 1
        order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", ["신규매수",
                                                                                                                            self.account_stock_dict[sCode]["주문용스크린번호"], self.account_num, 1, sCode, quantity, " ", self.realType.SENDTYPE['거래구분']['시장가'], ""])
        if order_success==0:
            print("성공")
            screen_num = self.account_stock_dict[sCode]["스크린번호"]
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall("SetRealReg(QString, QString, QString, QString)", screen_num, sCode, fids, "1")
        else:
            print("실패")
            del self.account_stock_dict[sCode]
    
    #매도 주문 전송
    def sendSellOrder(self, sCode):
        if sCode in self.account_stock_dict:
            quantity = self.account_stock_dict[sCode]['보유수량']
            order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                             ["신규매도", self.account_stock_dict[sCode]["주문용스크린번호"], self.account_num, 2, sCode, quantity, 0, self.realType.SENDTYPE['거래구분']['시장가'], ""])
            if order_success==0:
                print("성공")
                self.logging.logger.debug("종목코드: %s 매도" % (sCode))
                # print(self.self.account_stock_dict[sCode]["스크린번호"])
                self.dynamicCall("SetRealRemove(QString, QString)", self.account_stock_dict[sCode]["스크린번호"], sCode)
                del self.account_stock_dict[sCode]
                self.sellCount = 1
            else:
                print("실패")
        else:
            print("해당 주식 코드는 현재 계좌에 존재하지 않습니다.")

    #스크린 번호 Setting
    def screen_number_setting(self):
        self.cnt = 0
        for code in self.account_stock_dict.keys():
            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)
            if(self.cnt%50) == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)
            if(self.cnt%50) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)
            self.account_stock_dict[code].update({"스크린번호": str(self.screen_real_stock)})
            self.account_stock_dict[code].update({"주문용스크린번호": str(self.screen_meme_stock)})
            self.cnt += 1