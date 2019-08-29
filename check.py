import re
import pings
import urllib.request
from bs4 import BeautifulSoup
#引数に指定したipアドレスにpingし、疎通可能性を戻り値として返す関数
def Ping(ip):
    p = pings.Ping()
    res = p.ping(ip,times=4)
    print("-------------------------------------")
    res.print_messages()
    print("-------------------------------------")
    return res.is_reached()

def Phone_Check(ip):
    #webスクレイピングでデータ取得
    weburl = r"http://"+ip+"/CGI/Java/Serviceability?adapter=device.statistics.configuration"
    with urllib.request.urlopen(weburl) as source:
        html = source.read()
    #データを処理しやすく加工
    soup = BeautifulSoup(html,"lxml")
    settings = soup.find_all("b")
    txtdata = ""
    for data in settings:
        txtdata += " " + data.text.strip()
    #DHCPサーバを判定
    dhcp = re.search("DHCPサーバ 192.168.241.12",txtdata)
    if dhcp:
        print(dhcp.group())
        dhcpjudge = "OK"
    else:
        print("DHCPに問題があります")
        dhcpjudge = "NO"
    #TFTPサーバ1を判定
    tftp1 = re.search("TFTPサーバ1 192.168.241.35",txtdata)
    if tftp1:
        print(tftp1.group())
        tftp1judge = "OK"
    else:
        print("TFTP1に問題があります")
        tftp1judge = "NO"
    #CUCMサーバ1を判定
    cucm1 = re.search("CUCMサーバ1 192.168.241.34  Active",txtdata)
    #CUCMサーバ2を判定
    cucm2 = re.search("CUCMサーバ2 192.168.241.35  Standby",txtdata)
    if cucm1:
        print(cucm1.group())
        cucm1judge = "OK"
    else:
        cucm1judge = "NO"
    if cucm2:
        print(cucm2.group())
        cucm2judge = "OK"
    else:
        cucm2judge = "NO"
    if not (cucm1 and cucm2):
        print("正しいCUCMに登録されていません")
    #TVS判定
    tvs = re.search("TVS",txtdata)
    if tvs:
        print("ITLファイルインストール済")
        tvsjudge = "ITLファイルインストール済"
    else:
        print("ITLファイルがインストールされていません")
        tvsjudge = "ITLファイル未インストール"
    if dhcp and tftp1 and cucm1 and cucm2 and tvs:
        print("機器" + ip + "は正常に作動しています\n\n")
    else:
        print("機器" + ip + "に異常があります\n\n")
    return [dhcpjudge,tftp1judge,cucm1judge,cucm2judge,tvsjudge]
