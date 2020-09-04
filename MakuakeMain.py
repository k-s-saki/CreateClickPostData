import os

import csv
import tkinter
from tkinter import filedialog
import jaconv
import re

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
print("desktop => " + desktop)

'''
filetype = [("CSV", "*.csv"), ("すべて", "*")]
file_to_read = tkinter.filedialog.askopenfilename(filetypes=filetype, initialdir=desktop,title="マクアケのファイル")
print("read file => " + file_to_read)
if not file_to_read:
    print("canceled.")
    exit(-1)
'''

#file_to_read= desktop+"\makuake受発注\iu_air_mask_20200806_143402.csv"
file_to_read= desktop+"\makuake受発注\iu_air_mask_20200806_143402_adr.csv"

"""
def ProcessMatchStr():
    return "ALL"
def ProcessMatch(csv_line):
    return True

def ProcessMatchStr():
    return "8上_白細"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_white_1):
        return HosoFutoStr(csv_line[c_tem])=="細"
    return False

def ProcessMatchStr():
    return "8上_白太"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_white_1):
        return HosoFutoStr(csv_line[c_tem])=="太"
    return False

def ProcessMatchStr():
    return "8上_橙細"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_orange_1):
        return HosoFutoStr(csv_line[c_tem])=="細"
    return False

def ProcessMatchStr():
    return "8上_橙太"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_orange_1):
        return HosoFutoStr(csv_line[c_tem])=="太"
    return False

def ProcessMatchStr():
    return "8下_白細"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_white_2):
        return HosoFutoStr(csv_line[c_tem])=="細"
    return False

def ProcessMatchStr():
    return "8下_白太"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_white_2):
        return HosoFutoStr(csv_line[c_tem])=="太"
    return False

def ProcessMatchStr():
    return "8下_橙細"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_orange_2):
        return HosoFutoStr(csv_line[c_tem])=="細"
    return False
"""

def ProcessMatchStr():
    return "9上_Pack"
def ProcessMatch(csv_line):
    if csv_line[c_retid] == str(retid_dantai12):
        return True
    return False


#ファイルのパスと拡張子など
base_pair = os.path.split(file_to_read)
ext = os.path.splitext(base_pair[1])

# click_postラベルの最大数
MAX_CLICKPOST_LABEL=20
click_post_page=1

def GetClickPostFilename(nPage):
    # click_post用のファイル名
    click_post_fn = os.path.join(
        base_pair[0],
        ext[0] + "_clickpost_" + ProcessMatchStr() +"P"+str(nPage)+ ext[1])
    print("clickpost file => " + click_post_fn)
    return click_post_fn


#order用のファイル名
order_fn = os.path.join(
    base_pair[0],
    ext[0] + "_order_"+ProcessMatchStr()+".txt")
print("order file => " + order_fn)


#文字列をclickpostのアドレス(4行)にあわせて分割
def SplitClickPostAddr4(oAddrList, aStr):
    '''
    click postのアドレスは 1行20文字以内なので、
    4行20文字以内に行分割する (行分割できない場合はエラー終了)
    自動で行分割すると数字の途中で行分割されることもあるため、
    あらかじめ、分けたい箇所には、@を付けておく

    :param oAddrList: 住所リスト（出力）
    :param aStr: 住所文字列（入力）
    :return:
    '''
    # アドレス分割の正規表現
    address_pattern = "(...??[都道府県])((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山" \
                      "|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下>松|岩国|田川|大村|宮古|富良野|別府|佐伯|黒部|小諸|塩尻|玉野|周南)市|" \
                      "(?:余市|高市|[^市]{2,3}?)郡(?:玉村|大町|.{1,5}?)[町村]|(?:.{1,4}市)?[^町]{1,4}?区|.{1,7}?[市町村])(.+)"

    #すべて全角に変換
    #zStr= jaconv.h2z( aStr, kana=True, digit=True, ascii=True)
    zStr= aStr

    # 市区町村をリストに分解する(1行20文字以内で4行以内）
    adr_list=re.split(address_pattern, zStr)
    s=["","","",""]
    idx=0
    for adr_i in adr_list:
        for adr in adr_i.split('@'):
            if idx<=3:
                if len(s[idx])+len(adr)<=20:
                    s[idx]+=adr
                else:
                   idx+=1
                   s[idx]+=adr
            else:
                print("ClickPost 住所分割エラー："+ aStr)
                exit(-1)
    oAddrList.extend(s)

# bar
bar= "----------------------------------------------------------------------\n"

# MakuakeのCSVカラムの定義
c_id=0
c_userid=2
c_zip=13
c_full_addr=14
c_name=4
c_retid=8  #リターンID
c_date=1
c_status=12

retid_orange_1=87160
retid_orange_2=87161
retid_white_1=87159
retid_white_2=87162
retid_dantai12=87163

c_tem=17 #テンプルの選択


def HosoFutoStr(c_tem_str):
    if ("細" in c_tem_str):
        return "細"
    elif ("太" in c_tem_str):
        return "太"
    else:
        return "？？？"

def PrepareClickPostFile(nPage):
    file =open(GetClickPostFilename(nPage), mode="w", encoding="shift_jis", newline="")
    # クリックポストのヘッダー
    writer = csv.writer(file)
    writer.writerow([
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称",
        "お届け先住所1行目", "お届け先住所2行目", "お届け先住所3行目", "お届け先住所4行目",
        "内容品"])
    return file,writer


#ファイルの読み書き
with open(file_to_read, mode="r", encoding="shift_jis") as rf,\
     open(order_fn, mode="w", encoding="shift_jis", newline="") as order_wf:

    # STORES.JP の CSVを読み込む
    reader = csv.reader(rf)
    # 読み取りファイル　ヘッダー行を飛ばす
    next(reader)
    match_count=0

    click_post_page=0
    click_post_count=0
    click_post_wf, click_post_writer= PrepareClickPostFile(click_post_page)
    prev_id = ""

    print("抽出="+ProcessMatchStr())
    order_wf.write("オーダーリスト"+ProcessMatchStr()+"\n")

    # 読み取りループ
    for line in reader:
        # print("debug:"+','.join(line))
        print("check id="+line[c_id])

        if ProcessMatch(line):
            # 特定のデータのみ
            match_count+=1
            zip = line[c_zip]
            name = line[c_name]

            #Click Post用にアドレスを分割する
            adr = []
            SplitClickPostAddr4(adr, line[c_full_addr])
            adr.extend(["", "", "", ""])

            memo = "マスク-"+ProcessMatchStr()+ "." + line[c_id]  # order No.

            # click post write
            click_post_count+=1
            if click_post_count > MAX_CLICKPOST_LABEL:
                click_post_count=1
                click_post_page+=1
                click_post_wf.close()
                click_post_wf, click_post_writer = PrepareClickPostFile(click_post_page)

            click_post_writer.writerow([zip, name, '様', adr[0], adr[1], adr[2], adr[3], memo])


            # debug
            print("発送先データ:"+name +"様 〒" +zip )
            print("分割前:"+line[c_full_addr])
            print("(0)"+adr[0])
            print("(1)"+adr[1])
            print("(2)"+adr[2])
            print("(3)"+adr[3])
            print(memo)

            #オーダーNo、住所などの記入
            order_wf.write(bar)
            order_wf.write("#" + str(match_count) + " Order No." + line[c_id] + " Status=" + line[c_status] + " DATE=" + line[c_date] + "\n")
            order_wf.write("〒" +zip+"\n")
            order_wf.write("C:"+line[c_full_addr]+"\n")
            order_wf.write(adr[0]+" || "+adr[1]+"\n")
            order_wf.write(adr[2]+" || "+adr[3]+"\n")
            order_wf.write(name +"様 \n")
            order_wf.write(memo+"\n")

            #オーダーアイテムの記入
            """
            order_wf.write(line[8]+"\n")
            order_wf.write(line[9]+"\n")
            order_wf.write("個数:" + line[10]+"\n")
            order_wf.write("\n")
            """

    order_wf.write(bar)
    order_wf.write("抽出数=" + str(match_count))
    click_post_wf.close()
