import os

import csv
import tkinter
from tkinter import filedialog
import re
from enum import Enum, unique

@unique
class WorkDirConfig(Enum):
    # STD_DIALOG= 0
    NETWORK = 1
    LOCAL_DESKTOP = 2
    # SERVER = 3

""" User Configuration """

# ワークディレクトリの選択
WorkDir = WorkDirConfig(WorkDirConfig.NETWORK)


""" ------------ """


def GetWorkDir():
    if WorkDir == WorkDirConfig.NETWORK:
        return r'\\SVRZ230\L20\STORES_JP\WORK'   # <---- important  r
    elif WorkDir == WorkDirConfig.LOCAL_DESKTOP:
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        return desktop


filetype = [("CSV", "*.csv"), ("すべて", "*")]
file_to_read = tkinter.filedialog.askopenfilename(filetypes=filetype, initialdir=GetWorkDir())
print("read file => " + file_to_read)
if not file_to_read:
    print("canceled.")
    exit(-1)

# click_post用のファイル名
base_pair = os.path.split(file_to_read)
ext = os.path.splitext(base_pair[1])
click_post_fn = os.path.join(
    base_pair[0],
    ext[0] + "_click_post" + ext[1])
print("click_post file => " + click_post_fn)

# order用のファイル名
order_fn = os.path.join(
    base_pair[0],
    ext[0] + "_order.txt")
print("order file => " + order_fn)


"""
def str_to_addr(oAddrList, aStr):
    '''
    click postのアドレスは 1行20文字以内なので、
    20文字以上だったら分割する
    :param oAddrList: 住所リスト（出力）
    :param aStr: 住所文字列（入力）
    :return:
    '''
    if len(aStr) <= 20:
        oAddrList.append(aStr)
    else:
        oAddrList.extend([aStr[:20], aStr[20:]])
"""


def split_addr_for_clickpost(oAddrList, aStr):
    """
    文字列をclickpostのアドレス(4行)にあわせて分割

    click postのアドレスは 1行20文字以内なので、
    4行20文字以内に行分割する (行分割できない場合はエラー終了)
    自動で行分割すると数字の途中で行分割されることもあるため、
    あらかじめ、分けたい箇所には、@を付けておく

    :param oAddrList: 住所リスト（出力）
    :param aStr: 住所文字列（入力）
    :return: None
    """

    # アドレス分割の正規表現
    address_pattern = "(...??[都道府県])((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山" \
                      "|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下>松|岩国|田川|大村|宮古|富良野|別府|佐伯|黒部|小諸|塩尻|玉野|周南)市|" \
                      "(?:余市|高市|[^市]{2,3}?)郡(?:玉村|大町|.{1,5}?)[町村]|(?:.{1,4}市)?[^町]{1,4}?区|.{1,7}?[市町村])(.+)"

    # すべて全角に変換
    # zStr= jaconv.h2z( aStr, kana=True, digit=True, ascii=True)
    zStr = aStr

    # 市区町村をリストに分解する(1行20文字以内で4行以内）
    adr_list=re.split(address_pattern, zStr)
    s=["","","",""]
    idx = 0
    for adr_i in adr_list:
        # 手作業で@を入れてあれば分割してみる
        for adr in adr_i.split('@'):
            if idx <= 3:
                if len(s[idx])+len(adr) <= 20:
                    # 足して20文字以内だったら、現在のインデックスに追加する
                    s[idx] += adr
                else:
                    # 足して20文字以上だったら、次のインデックスへ
                    idx += 1
                    s[idx] += adr
            else:
                print("ClickPost 住所分割エラー："+ aStr)
                exit(-1)
    oAddrList.extend(s)


# bar
bar_customer\
    = "=======================================================================\n"
bar_item\
    = "-----------------------------------------------------------------------\n"


#ファイルの読み書き
with open(file_to_read, mode="r", encoding="shift_jis") as rf,\
     open(click_post_fn, mode="w", encoding="shift_jis", newline="") as click_post_wf, \
     open(order_fn, mode="w", encoding="shift_jis", newline="") as order_wf:

    # STORES.JP の CSVを読み込む
    reader = csv.reader(rf)
    # 読み取りファイル　ヘッダー行を飛ばす
    next(reader)

    # クリックポストのヘッダー
    click_post_w = csv.writer(click_post_wf)
    click_post_w.writerow([
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称",
        "お届け先住所1行目", "お届け先住所2行目", "お届け先住所3行目", "お届け先住所4行目",
        "内容品"])
    prev_id = ""

    # 読み取りループ
    for line in reader:
        # print("debug:"+','.join(line))

        status = line[1]
        if status == "未発送":
            # 未発送の場合のみ処理をする
            if prev_id != line[0]:
                # 同じ注文IDではない場合・・送り先を記入
                prev_id = line[0]
                zip = line[32]
                name = line[30] + line[31]
                adr = []
                split_addr_for_clickpost(adr, line[33]+line[34])
                adr.extend(["", "", "", ""])
                c = "マスク SJ." + line[0]  # order No.
                click_post_w.writerow([zip, name, '様', adr[0], adr[1], adr[2], adr[3], c])
                print("発送先データ:" + ','.join(line))
                # オーダーNoの記入
                order_wf.write(bar_customer)
                order_wf.write("STORE.JP No."+line[0] +" Status="+line[1]+" DATE="+line[3]+"\n")

                # 住所の記入
                order_wf.write("\n")
                order_wf.write("送り先：" + "\n")
                order_wf.write("〒" + zip + "\n")
                order_wf.write(adr[0] + "\n")
                order_wf.write(adr[1] + "\n")
                order_wf.write(adr[2] + "\n")
                order_wf.write(adr[3] + "\n")
                order_wf.write(name + "様 \n")
                order_wf.write(bar_customer)
                order_wf.write("注文内容:" + "\n")

            # オーダーアイテムの記入
            order_wf.write(bar_item)
            order_wf.write(line[8]+"\n")
            order_wf.write(line[9]+"\n")
            order_wf.write("個数:" + line[10]+"\n")
            if line[44]:
                order_wf.write("メモ:" + line[44] + "\n")
            order_wf.write("\n")

