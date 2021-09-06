import requests
import configparser
import xml.etree.ElementTree as ET

class Data():
    CORP_CODE_URL ="http://api.seibro.or.kr/openapi/service/CorpSvc/getIssucoCustnoByNm"
    CORP_INFO_URL ="http://api.seibro.or.kr/openapi/service/CorpSvc/getIssucoBasicInfo"
    STOCK_DISTRIBUTION_URL="http://api.seibro.or.kr/openapi/service/CorpSvc/getStkDistributionShareholderStatus"

    def __init__(self):
        #api키 설정
        config =configparser.RawConfigParser()
        config.read("conf/config.ini")
        self.api_key =config["DATA"]["api_key"]
        if self.api_key is None:
            raise Exception("Need to api key")

    def get_corp_code(self,name=None):
        #회사명으로 기업코드찾기
        query_params = {"ServiceKey":self.api_key,"issucoNm":name,"numOfRows":str(5000)}
        request_url =self.CORP_CODE_URL+"?"
        for k, v in query_params.items():
            request_url = request_url+k+"="+v+"&"

        res=requests.get(request_url[:-1])
        root=ET.fromstring(res.text)
        from_tags = root.items("items")
        result = {}
        for items in from_tags:
            for item in items.iter("item"):
                if name in item.find('issucoNm').text.split():
                    result["issucoCustno"]=item.find('issucoCustno').text
                    result["issucoNm"]=item.find('issucoNm').text
        return result

    def get_corp_info(self,code=None):
        #기업개요 조회
        query_parms={"ServiceKey":self.api_key,"issucoCustno":code.replace("0","")}
        request_url = self.CORP_INFO_URL + "?"
        for k, v in query_params.items():
            request_url = request_url + k + "=" + v + "&"

        res=requests.get(request_url[:-1])
        root = ET.fromstring(res.text)
        from_tags = root.iter("item")
        result ={}
        for item in from_tags:
            result["apliDt"] = item.find("apliDt").text
            result["bizno"] = item.find("bizno").text
            result["ceoNm"] = item.find("ceoNm").text
            result["engCustNm"] = item.find("apliDt").text
            result["foundDt"] = item.find("foundDt").text
            result["homepAddr"] = item.find("homepAddr").text
            result["pval"] = item.find("pval").text
            result["totalStkcnt"] = item.find("totalStkcnt").text
            result[""] = item.find("").text





