#----------------------------------------------
# 2020-09-19  タイトルのズレと罫線のサイズ変更
# 2020-10-29  出力するテキストデータの場所を C:\Users\uenok\Nippo\ に変更
#----------------------------------------------

# 労働時間の報告
import os
import xlrd
import datetime

#----------
#  職員数を設定する ※間違うとエラーになる
#----------
staffvol= 15
#----------
# 報告しない人
#----------
reigai = {"山本智之": True, "竹内啓子": True, "村井直行": True, "國吉裕": True}

cnt: int
stcnt, cnt = 9, 0

#女性の行で区切る先頭の人
joseitop="西村益江"

#baseyyyymmdd.xlsm の場所
os.chdir('C:\\Users\\uenok\\Nippo')

#atennshonn
print("総労働時間通知.txtは閉じておいてください")
print("昨日までの実績ならば本日の日付をいれる")
workdate=input('yymmdd = ')

#文字列を日時に変換(strptime)し、日付だけ取り出し(.date())、そこから1日引いて前日の日付を求める(timelta(days=1))
backdate=datetime.datetime.strptime(workdate,'%y%m%d').date()-datetime.timedelta(days=1)

#年月日を個別に取り出し
header = ['From:     上野浩司 <uenokoji@freedom21.jp>',
          'To:       ukaigi@f21mail.org',
          'Subject: '+str(backdate.year)+"年"+str(backdate.month)+"月"+str(backdate.day)+"日 勤務実績",
          'X-TuruKame-KeitaiSend: 1\n',
          '',
          '時間:移動を含み、有休は含まず\n',
          '名前　　　　時間(正味) 公休 有休',
          '-' * 32]
print("ファイル読み込み中")

# ファイル名を作る
book = xlrd.open_workbook('base' + workdate + '.xlsm')
sheet = (book.sheet_by_name('Main'))

#総労働時間通知.txtをオープン
fw = open("総労働時間通知.txt", "w")

for h in header:
    fw.writelines(h + '\n')
#1名につき1つの隠れ行があるので2つ飛ばしでスタッフの行数分Loopする
while staffvol > cnt / 2: 
    sname = sheet.cell_value(stcnt + cnt, 1)
    if sname==joseitop:#女性と男性を1行空ける判断
        fw.writelines('\n')
    if reigai.get(sname, False):
        pass
#総労働時間、正味労働時間は0埋め3桁表示
    else:
        fw.writelines(sname + '　' * (6 - len(sname))) #フルネーム
        fw.writelines('' +'{:03}'.format(int(sheet.cell_value(stcnt + cnt, 23)))) #総労働時間
        fw.writelines(' (' +'{:03}'.format(int(sheet.cell_value(stcnt + cnt, 21)))) #正味労働時間
        fw.writelines(')   ' +'{0:2}'.format(int(sheet.cell_value(stcnt + cnt, 25)))) #公休
        fw.writelines('   ' +'{0:2}'.format(int(sheet.cell_value(stcnt + cnt, 29))) + '\n') #有休
    cnt += 2
fw.writelines('-' * 32 + '\n※小数点以下四捨五入\n以上です')
fw.close()

# ファイルオープン
os.chdir("C:\\Program Files (x86)\\Sakura")
os.system('sakura.exe C:\\Users\\uenok\\Nippo\\'+fw.name)
