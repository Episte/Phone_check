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
    #IP電話のwebページでwebスクレイピング
    weburl = r"http://"+ip+"/CGI/Java/Serviceability?adapter=device.statistics.configuration"
    try:
        with urllib.request.urlopen(weburl) as source:
            html = source.read()
        #データを処理しやすく加工。必要なデータはｂ囲まれているため、ｂを指定
        soup = BeautifulSoup(html,"lxml")
        settings = soup.find_all("b")
        txtdata = ""
        for data in settings:
            txtdata += " " + data.text.strip()
        #登録情報の判定処理
        reglist = ["MACアドレス ","DHCPサーバ ","TFTPサーバ1 ",\
        "TFTPサーバ2 ","CUCMサーバ1 ","CUCMサーバ2 ","TVS"]
        #戻り値を作成（UnboundLocalError対策）
        judgereg = []
        mac_cr = ""
        """
        取得したIP電話の情報から、reglistで用意した文字列を探す。あった場合は後ろに続く機器固有の
        情報をリストに格納。なかった場合は×を格納
        """
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
    except:
        print("WebアクセスがEnabledに設定されていません。CUCMで設定してください。")
