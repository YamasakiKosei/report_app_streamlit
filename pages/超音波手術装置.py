# コメント
'''
外装点検は、総合評価の合否に考慮されていない
Excelを変更したら、入力するセルの位置を確認
'''

# 初期設定
file_path = './点検報告書/超音波手術装置 点検報告書 org.xlsx' # 雛形Excelのパス
eval_ctg = ['電気的安全性点検']                             # 総合評価の評価項目



import streamlit as st
from openpyxl import load_workbook
from package import module

# インスタンス 初期化
module.stObject.instances = []



# タイトル
module.Header('超音波手術装置')

# 機器情報
info1, info2, info3, info4, info5, info6 = module.Info()

# 電気的安全性点検
eli1, eli2, eli3, eli4, eli5 = module.Eli()

# 外装点検
st.subheader('外装点検', help='総合評価には含まれません')
outi1 = module.stObject('外装点検', '外装点検１', '１．装置に汚れ、ひび割れ、破損がないか')
outi2 = module.stObject('外装点検', '外装点検２', '２．アクティブコネクタ、マイクロバイポーラコネクタ、対極板接続コネクタが清拭されているか')
outi3 = module.stObject('外装点検', '外装点検３', '３．フットスイッチ及びコードに汚れや破損がないか')
outi4 = module.stObject('外装点検', '外装点検４', '４．メス先電極及びコネクタ、対極板コード及びコネクタの破損がないか')
outi5 = module.stObject('外装点検', '外装点検５', '５．メス先電極ホルダ、ハンドコントロールスイッチの破損がないか')
outi6 = module.stObject('外装点検', '外装点検６', '６．電源コードの腐食や破損がないか')
module.Checkbox(outi1, outi2, outi3, outi4, outi5, outi6)
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
    sheet['B22'] = text_area.value
    # 総合評価
    sheet['H25'] = '合格' if not errorList else '不合格'

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