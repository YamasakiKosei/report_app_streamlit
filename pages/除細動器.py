# コメント
'''
外装点検は、総合評価の合否に考慮されていない
Excelを変更したら、入力するセルの位置を確認
'''

# 初期設定
file_path = './点検報告書/除細動器 点検報告書 org.xlsx' # 雛形Excelのパス
eval_ctg = ['電気的安全性点検', '機能点検']           # 総合評価の評価項目



import streamlit as st
from openpyxl import load_workbook
from package import module

# インスタンス 初期化
module.stObject.instances = []



# タイトル
module.Header('除細動器')

# 機器情報
info1, info2, info3, info4, info5, info6 = module.Info()

# 電気的安全性点検
eli1, eli2, eli3, eli4, eli5 = module.Eli()

# 外装点検
st.subheader('外装点検', help='総合評価には含まれません')
outi1 = module.stObject('外装点検', '外装点検１', '１．装置に汚れ、ひび割れ、破損がないか')
outi2 = module.stObject('外装点検', '外装点検２', '２．操作パネルの傷、はく離がないか')
outi3 = module.stObject('外装点検', '外装点検３', '３．部品のゆるみ、ねじ、ナットのゆるみがないか')
outi4 = module.stObject('外装点検', '外装点検４', '４．電源コードの破損がないか')
module.Checkbox(outi1, outi2, outi3, outi4)
st.divider()

# 機能点検
st.subheader('機能点検')
fni1 = module.stObject('機能点検', '機能点検１', '１．現在時刻の確認')
fni2 = module.stObject('機能点検', '機能点検２', '２．電源を入れた際にセルフチェック機能が働くか')
fni3 = module.stObject('機能点検', '機能点検３', '３．動作中・異常発生時のランプが点灯・点滅するか')
module.Checkbox(fni1, fni2, fni3)
st.divider()

# 総合評価
errorList = module.Evaluation(eval_ctg)

# 備考
text_area = module.Remarks()

# ダウンロード
module.FileName() # 現在のファイル名 表示

# 作成ボタン
module.CreateButton()

# 個別設定
def unique(sheet):
    # 備考
    sheet['B23'] = text_area.value
    # 総合評価
    sheet['H26'] = '合格' if not errorList else '不合格'

# ダウンロードボタン
if st.session_state['ダウンロードボタン']:
    # Excelに入力
    wb = load_workbook(file_path)
    sheet = wb['Original']
    module.WriteCommon(sheet) # 共通項目
    unique(sheet) # 個別設定
    data = module.Save(wb) # 保存
    module.DownloadButton(data)

# プログレスバー
module.ProgressBar()