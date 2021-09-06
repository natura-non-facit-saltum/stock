import configparser
import win32com.client
import pythoncom
from datetime import datetime
import time

class XASession:
    login_state = 0

    def OnLogin(self,code,msg):
        if code =="0000":
            print(code,msg)
            XASession.login_state = 1
        else:
            print(code,msg)

    def OnDisconnect(self):
        print("Session disconnected")
        XASession.login_state= 0

class XAQuery:
    RES_PATH = "C:\\eBEST\\XingAPI\\Res\\"
    tr_run_state = 0

    def OnReceiveData(self,code):
        print("OnReceiveData",code)
        XAQuery.tr_run_state= 1

    def OnReceiveMessage(self,error,code,message):
        print("OnReceiveMessage",error,code,message)



class EBest:
    QUERY_LIMIT_10MIN = 200
    LIMIT_SECONDS = 600

    def __init__(self,mode=None):
        if mode not in["PROD","DEMO"]:
            raise Exception("Need to run_mode(PROD or DEMO)")

        run_mode ="EBEST_"+mode
        config = configparser.ConfigParser()
        config.read("conf/config.ini")
        self.user = config[run_mode]['user']
        self.passwd = config[run_mode]['password']
        self.cert_passwd = config[run_mode]['cert_passwd']
        self.host = config[run_mode]['host']
        self.port = config[run_mode]['port']
        self.account = config[run_mode]['account']

        # 세션 com 설정
        self.xa_session_client = win32com.client.DispatchWithEvents('XA_Session.XASession',XASession)
        self.query_cnt=[]

    def login(self):

        # 세션 요청
        self.xa_session_client.ConnectServer(self.host,self.port)
        self.xa_session_client.Login(self.user,self.passwd,self.cert_passwd,0,0)
        while XASession.login_state == 0:
            pythoncom.PumpWaitingMessages()

    def logout(self):
        result = self.xa_session_client.Logout()
        if result:
            XASession.login_state = 0
            self.xa_session_client.DisconnectServer()

    def _execute_query(self,res,in_block_name,out_block_name,*out_fields,**set_fields):

        # 10분200회 제한
        time.sleep(1)
        print("current query cnt:",len(self.query_cnt))
        print(res,in_block_name,out_block_name)
        while len(self.query_cnt) >= EBest.QUERY_LIMIT_10MIN:
            time.sleep(1)
            print("waitng for execute query... current query cnt:",len(self.query_cnt))
            self.query_cnt=list(filter(lambda x:(datetime.today()-x).total_seconds()<EBest.LIMIT_SECONDS,self.query_cnt))

        # 쿼리 com ,res 설정
        xa_query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery",XAQuery)
        xa_query.LoadFromResFile(XAQuery.RES_PATH+res+".res")

        # in_block_name 셋팅 및 쿼리요청
        for key, value in set_fields.items():
            xa_query.SetFieldData(in_block_name,key,0,value)
        errorCode = xa_query.Request(0)

        #요청 후 대기
        waiting_cnt = 0
        while XAQuery.tr_run_state==0:
            waiting_cnt +=1
            if waiting_cnt%100000== 0:
                print("Waiting....",self.xa_session_client.GetLastError())
            pythoncom.PumpWaitingMessages()

        #결과블록
        result=[]
        count = xa_query.GetBlockCount(out_block_name)

        for i in range(count):
            item = {}
            for field in out_fields:
                value = xa_query.GetFieldData(out_block_name,field,i)
                item[field] = value
            result.append(item)

        #제약시간 체크
        XAQuery.tr_run_state= 0
        self.query_cnt.append(datetime.today())

        #영문 필드명을 한글 필드명으로 변환
        for item in result:
            for field in list(item.keys()):
                if getattr(Field,res,None):
                    res_field = getattr(Field,res,None)
                    if out_block_name in res_field:
                        field_hname = res_field[out_block_name]
                        if field in field_hname:
                            item[field_hname[field]] = item[field]
                            item.pop(field)
        return result

    def get_code_list(self,market=None):
        #t8486 코스피,코스닥의 종목 리스트를 가져온다.

        if market !="ALL" and market !="KOSPI" and market !="KOSDAQ":
            raise Exception("Need tomarket param(ALL,KOSPI,KOSDAQ)")

        market_code={"ALL":"0","KOSPI":"1","KOSDAQ":"2"}
        in_params ={"gubun":market_code[market]}
        out_parms =["hname","shcode","expcode","etfgubun","recprice","gubun","spac_gubun"]
        result =self._execute_query("t8436","t8436InBlock","t8436OutBlock",*out_parms,**in_params)

        return result

    def get_stock_price_by_code(self,code=None,cnt="1"):
        #t1305 현재 날짜를 기준으로 cnt만큼 전일의 데이터를 가져온다.

        in_params={"shcode":code,"dwmcode":"1","date":"","idx":"","cnt":cnt}
        out_parms=["date","open","high","low","close","sign","change","diff","volume","diff_vol","chdegree","sojinrate","changerate","fpvolume",\
                   "covolume","value","ppvolume","o_sign","o_change","o_diff","h_sign","h_change","h_diff","l_sign","l_change","l_diff","marketcap"]

        result = self._execute_query("t1305","t1305Inblock","t1305OutBlock1",*out_parms,**in_params)

        for item in result:
            item["code"]=code

        return result

    def get_credit_trend_by_code(self,code=None,date=None):
        #t1921신용거래 동향

        in_params = {"gubun": "0","shcode": code,"date":date,"idx":"0"}
        out_parms = ["mmdate","close","sign","jchange","diff","nvolume","svolume","jvolume","price","change","gyrate","jkrate","shcode"]
        result = self._execute_query("t1921", "t1921InBlock", "t1921OutBlock1", *out_parms, **in_params)

        return result

    def get_agent_trend_by_code(self,code=None,fromdt=None,todt=None):
        #t1717 외인기관별 종목별 동향

        in_params = {"gubun": "0","shcode": code,"fromdt":fromdt,"todt":todt}
        out_parms = ["date","close","sign","change","diff","volume","tjj000_vol","tjj001_vol","tjj002_vol","tjj003_vol","tjj004_vol","tjj005_vol",\
                     "tjj006_vol","tjj00_vol","tjj008_vol","tjj009_vol","tjj010_vol","tjj011_vol","tjj018_vol","tjj016_vol","tjj017_vol","tjj000_dan",\
                     "tjj001_dan","tjj002_dan","tjj003_dan","tjj004_dan","tjj005_dan","tjj006_dan","tjj007_dan","tjj008_dan","tjj009_dan","tjj010_dan",\
                     "tjj011_dan","tjj018_dan""tjj016_dan""tjj017_dan"]
        result = self._execute_query("t1717", "t1717InBlock", "t1717OutBlock", *out_parms, **in_params)

        for item in result:
            item["code"]=code
        return result

    def get_short_trend_by_code(self,code=None,sdate=None,edate=None):
        #t1927 공매도 일별추이신용거래 동향

        in_params = {"date":sdate,"sdate":sdate,"edate":edate,"shcode":code}
        out_parms = ["date","price","sign","change","diff","volume","value","gm_vo","gm_va","gm_per","gm_avg","gm_vo_sum"]
        result = self._execute_query("t1927", "t1927InBlock", "t1927OutBlock1", *out_parms, **in_params)

        for item in result:
            item["code"] = code
        return result


class Field:

    t1305 = {
        "t1305OutBlock1": {
            "date": "날짜",
            "open": "시가",
            "high": "전일대비구분",
            "low" : "저가",
            "close" : "종가",
            "sign" : "전일대비구분",
            "change": "전일대비",
            "diff" : "등락율",
            "vloume": "누적거래량",
            "diff_vol" : "거래증가율",
            "chdegree" : "체결강도",
            "sojinrate" : "소진율",
            "changerate": "회전율",
            "fpvolume" : "외인순매수",
            "covolume" : "기관순매수",
            "shcode" :"종목코드",
            "value" : "누적거래대금(백만)",
            "ppvolume": "개인순매수",
            "o_sign" : "시가대비구분",
            "o_change" :"시가대비",
            "o_diff" :"시가기준등락율",
            "h_sign" :"고가대비구분",
            "h_change": "고가대비",
            "h_diff":"고가기준등락율",
            "l_sign":"저가대비구분",
            "l_change":"저가대비",
            "l_diff":"저가기준등락율",
            "marketcap":"시가총액(백만)"

        }
    }


    t1921 = {
        "t1921OutBlock1": {
            "mmdate":"날짜",
            "close":"종가",
            "sign":"전일대비구분",
            "jchange":"전일대비",
            "diff":"등락율",
            "nvolume":"신규",
            "svolume":"상환",
            "jvolume":"잔고",
            "price":"금액",
            "change":"대비",
            "gyrate":"공여율",
            "jkrate":"잔고율",
            "shcode":"종목코드"
        }
    }

    t1717 = {
        "t1717OutBlock": {
            "date" :"일자",
            "close":"종가",
            "sign" :"전일대비구분",
            "change":"전일대비",
            "diff":"등락율",
            "volume":"누적거래량",
            "tjj000_vol":"사모펀드(순매수량)",
            "tjj001_vol":"증권(순매수량)",
            "tjj002_vol":"보험(순매수량)",
            "tjj003_vol":"투신(순매수량)",
            "tjj004_vol":"은행(순매수량)",
            "tjj005_vol":"종금(순매수량)",
            "tjj006_vol":"기금(순매수량)",
            "tjj007_vol":"기타법인(순매수량)",
            "tjj008_vol":"개인(순매수량)",
            "tjj009_vol":"등록외국인(순매수량)",
            "tjj010_vol":"미등록외국인(순매수량)",
            "tjj011_vol":"국가외(순매수량)",
            "tjj018_vol":"기관(순매수량)",
            "tjj016_vol":"외인계(순매수량)",
            "tjj017_vol":"기타계(순매수량)",
            "tjj000_dan":"사모펀드(단가)",
            "tjj001_dan":"증권(단가)",
            "tjj002_dan":"보험(단가)",
            "tjj003_dan":"투신(단가)",
            "tjj004_dan":"은행(단가)",
            "tjj005_dan":"종금(단가)",
            "tjj006_dan":"기금(단가)",
            "tjj007_dan":"기타법인(단가)",
            "tjj008_dan":"개인(단가)",
            "tjj009_dan":"등록외국인(단가)",
            "tjj010_dan":"미등록외국인(단가)",
            "tjj011_dan":"국가외(단가)",
            "tjj018_dan":"기관(단가)",
            "tjj016_dan":"외인계(단가)",
            "tjj017_dan":"기타계(단가)"

        }
    }

    t1927 = {
        "t1927OutBlock1": {
            "date":"일자",
            "price":"현재가",
            "sign":"전일대비구분",
            "change":"전일대비",
            "diff":"등락율",
            "volume":"거래량",
            "value":"거래대금",
            "gm_vo":"공매도수량",
            "gm_va":"공매도대금",
            "gm_per":"공매도비중",
            "gm_avg":"평균공매도단가",
            "gm_vo_sum":"누적공매도수량",
            "gm_vo1":"업틱룰적용공매도수량",
            "gm_va1":"업틱룰적용공매도대금",
            "gm_vo2":"업틱룰예외공매도수량",
            "gm_va2":"업틱룰예외공매도대금"
        }
    }

    t8436 = {
        "t8436OutBlock": {
            "hname": "종목명",
            "shcode":"단축코드",
            "expcode":"확장코드",
            "etfgubun":"ETF구분(1:ETF,2:ETN)",
            "uplmtprice":"상한가",
            "dnlmtprice":"하한가",
            "jnilclose":"전일가",
            "recprice":"주문수량단위",
            "gubun":"구분(1:코스피,2:코스닥)",
            "bu12gubun":"증권그룹",
            "spac_gubun":"기업인수목적회사여부(Y/N)",
            "filter":"(미사용)"
        }
    }

    FOCCQ33600 = {
        "FOCCQ33600OutBlock1": {
            "RecCnt":"레코드갯수",
            "AcntNo":"계좌번호",
            "Pwd":"비밀번호",
            "QrySrtDt":"조회시작일",
            "QryEndDt":"조회종료일",
            "TermTp":"기간구분"
        },
        "FOCCQ33600OutBlock2": {
            "RecCnt":"레코드갯수",
            "AcntNm":"계좌명",
            "BnsctrAmt":"매매약정금액",
            "MnyinAmt":"입금금액",
            "MnyoutAmt":"출금금액",
            "InvstAvrbalPramt":"투자원금평잔금액",
            "InvstPlAmt":"투자손익금액",
            "invstErnrat":"투자수익률"
        },
        "FOCCQ33600OutBlock3": {
            "BaseDt":"기준일",
            "FdEvalAmt":"기초평가금액",
            "EotEvalAmt":"기말평가금액",
            "InvvstAvrbalPramt":"투자원금평잔금액",
            "BnsctrAmt":"매매약정금액",
            "MnyinSecinAmt":"입금고액",
            "MnyoutSecoutAmt":"출금고액",
            "EvalPnlAmt":"평가손익금액",
            "TermErnrat":"기간수익률",
            "idx":"지수"
        }
    }