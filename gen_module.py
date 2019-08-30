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
    try:
        with urllib.request.urlopen(weburl) as source:
            html = source.read()
    except:
        print("WebアクセスがEnabledに設定されていません。CUCMで設定してください。")
    else:
        #データを処理しやすく加工
        soup = BeautifulSoup(html,"lxml")
        settings = soup.find_all("b")
        txtdata = ""
        for data in settings:
            txtdata += " " + data.text.strip()
        #登録情報の判定処理
        reglist = ["MACアドレス ","DHCPサーバ ","TFTPサーバ1 ",\
        "TFTPサーバ2 ","CUCMサーバ1 ","CUCMサーバ2 ","TVS"]
        judgereg = []
        mac_cr = ""
        for reg in reglist:
            if reg == "TVS":
                if re.search(reg,txtdata):
                    print("ITLファイルインストール済")
                    judgereg.append("○")
                else:
                    print("ITLファイルがインストールされていません")
                    judgereg.append("×")
            elif reg == "MACアドレス ":
                mac = re.search(reg + "\w{12}",txtdata)
                if mac:
                    print(mac.group())
                    mac_cr = mac.group().replace(reg,"")
                else:
                    print(reg + "×")
            else:
                reg_ok = re.search(reg + "(\d{1,3}.){3}.\d{1,3}",txtdata)
                if reg_ok:
                    print(reg_ok.group())
                    reg_cr = reg_ok.group()
                    judgereg.append(reg_cr.replace(reg,""))
                else:
                    print(reg + "×")
                    judgereg.append("×")

        return mac_cr,judgereg


'''
    #DHCPサーバを判定
    dhcp = re.search("DHCPサーバ 192.168.241.\d{1,3}",txtdata)
    if dhcp:
        print(dhcp.group())
        dhcpjudge = dhcp.group().replace("DHCPサーバ ","")
    else:
        print("DHCPに問題があります")
        dhcpjudge = "×"
    #TFTPサーバ1を判定
    tftp1 = re.search("TFTPサーバ1 192.168.241.\d{1,3}",txtdata)
    if tftp1:
        print(tftp1.group())
        tftp1judge = tftp1.group().replace("TFTPサーバ1 ","")
    else:
        print("TFTP1に問題があります")
        tftp1judge = "×"
    #TFTPサーバ2を判定
    tftp2 = re.search("TFTPサーバ2 192.168.241.\d{1,3}",txtdata)
    if tftp2:
        print(tftp2.group())
        tftp2judge = tftp2.group().replace("TFTPサーバ2 ","")
    else:
        print("TFTP2に問題があります")
        tftp2judge = "×"
    #CUCMサーバ1を判定
    cucm1 = re.search("CUCMサーバ1 192.168.241.\d{1,3}",txtdata)
    #CUCMサーバ2を判定
    cucm2 = re.search("CUCMサーバ2 192.168.241.\d{1,3}",txtdata)
    if cucm1:
        print(cucm1.group())
        cucm1judge = cucm1.group().replace("CUCMサーバ1 ","")
    else:
        cucm1judge = "×"
    if cucm2:
        print(cucm2.group())
        cucm2judge = cucm2.group().replace("CUCMサーバ2 ","")
    else:
        cucm2judge = "×"
    #TVS判定
    tvs = re.search("TVS",txtdata)
    if tvs:
        print("ITLファイルインストール済")
        tvsjudge = "○"
    else:
        print("ITLファイルがインストールされていません")
        tvsjudge = "×"

    if dhcp and tftp1 and cucm1 and cucm2 and tvs:
        print("機器" + ip + "は正常に作動しています\n\n")
    else:
        print("機器" + ip + "に異常があります\n\n")

    return [dhcpjudge,tftp1judge,tftp2judge,cucm1judge,cucm2judge,tvsjudge]
'''
