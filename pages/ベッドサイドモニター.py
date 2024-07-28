# コメント
'''
外装点検は、総合評価の合否に考慮されていない
Excelを変更したら、入力するセルの位置を確認
'''

file_path = './点検報告書/ベッドサイドモニター 点検報告書 org.xlsx' # 雛形Excelのパス
sheet_name = 'Original'                                          # 雛形Excelのパス
eval_ctg = ['電気的安全性点検', '機能点検', '性能点検']             # 総合評価の評価項目
ctg = ['電気的安全性点検', '外装点検', '機能点検', '性能点検']       # 点検項目



import streamlit as st
import datetime
from openpyxl import load_workbook
from io import BytesIO
import time



# クラス
class stObject:
    instances = [] # クラス変数
    
    def __init__(self, category, name, label, value='', bool=''):
        self.category = category
        self.name = name
        self.label = label
        self.value = value
        self.bool = bool
        stObject.instances.append(self)
    
    @classmethod
    def getInstances(cls): return cls.instances # 作成したインスタンスをすべて取得

# インスタンス
info1 = stObject('機器情報', '機種名', '１．機種名')
info2 = stObject('機器情報', '製造番号', '２．製造番号')
info3 = stObject('機器情報', '製造販売業者', '３．製造販売業者')
info4 = stObject('機器情報', '機器管理番号', '４．機器管理番号')
info5 = stObject('機器情報', '購入年月日', '５．購入年月日')
info6 = stObject('機器情報', '実施年月日', '６．実施年月日')

eli1 = stObject('電気的安全性点検', '接触電流（NC）', '正常状態（NC）：100μA以下', '   ー', 'ー')
eli2 = stObject('電気的安全性点検', '接触電流（SFC）', '単一故障状態（SFC）：500μA以下', '   ー', 'ー')
eli3 = stObject('電気的安全性点検', '接地漏れ電流（NC）', '正常状態（NC）：5,000μA以下', '   ー', 'ー')
eli4 = stObject('電気的安全性点検', '接地漏れ電流（SFC）', '単一故障状態（SFC）：10,000μA以下', '   ー', 'ー')
eli5 = stObject('電気的安全性点検', '接地線抵抗', '0.2Ω以下', '   ー', 'ー')

outi1 = stObject('外装点検', '外装点検１', '１．装置に汚れ、ひび割れ、破損がないか')
outi2 = stObject('外装点検', '外装点検２', '２．操作パネルの傷、はく離がないか')
outi3 = stObject('外装点検', '外装点検３', '３．部品のゆるみ、ねじ、ナットのゆるみがないか')
outi4 = stObject('外装点検', '外装点検４', '４．電源コードの破損がないか')

fni1 = stObject('機能点検', '機能点検１', '１．現在時刻の確認')
fni2 = stObject('機能点検', '機能点検２', '２．電源を入れた際にセルフチェック機能が働くか')
fni3 = stObject('機能点検', '機能点検３', '３．動作中・異常発生時のランプが点灯・点滅するか')

peri1 = stObject('性能点検', '酸素飽和度', '１．酸素飽和度測定精度（93 ～ 100 %）')
peri2 = stObject('性能点検', '脈拍数', '２．脈拍数測定精度（69 ～ 75 拍）')
peri3 = stObject('性能点検', '心電図', '３．入力した波形が正常に出力されるか')
peri4 = stObject('性能点検', '体温部', '４．体温測定精度', '   ー', 'ー')
peri5 = stObject('性能点検', '血圧部（上限）', '５．最高血圧警報測定(上限値)')
peri6 = stObject('性能点検', '血圧部（下限）', '６．最高血圧警報測定(下限値)')



# Web
st.title('ベッドサイドモニター 点検報告書')
st.caption('点検報告書を作成し、Excel形式で保存できます')
col1, col2 = st.columns([1,2])
with col1:
    col1_1, col1_2, col1_3 = st.columns([4,1,8])
    with col1_1:
        st.page_link('Home.py', label='ホーム')
    with col1_2:
        st.write('**>**')
    with col1_3:
        st.page_link('pages/ベッドサイドモニター.py', label='ベッドサイドモニター')
st.divider()



# 機器情報
st.header('機器情報')
col1, col2 = st.columns(2)
with col1:
    info1.value = st.text_input(info1.label) # 機種名
    info3.value = st.text_input(info3.label) # 製造販売業者
with col2:
    info2.value = st.text_input(info2.label) # 製造番号
    def update_default():
        st.session_state['デフォルト'] = st.session_state[info4.name] + ' 点検報告書' # デフォルトを更新
        if not st.session_state['トグル']: # トグルがOFFだったら
            st.session_state['ファイル名'] = st.session_state['デフォルト'] # 表示するファイル名をデフォルトにする
            st.session_state['新ファイル名'] = st.session_state['デフォルト'] # 新ファイル名をデフォルトにする
    info4.value = st.text_input(info4.label, key=info4.name, on_change=update_default) # 機器管理番号
col1, col2 = st.columns(2)
with col1:
    info5.value = st.date_input(info5.label, value=datetime.date(2000, 1, 1), min_value=datetime.date(1900, 1, 1), max_value=datetime.date(2099, 12, 31)) # 購入年月日
with col2:
    info6.value = st.date_input(info6.label) # 実施年月日
st.divider()



# 点検
st.header('点検')

# 電気的安全性点検
toggle1 = st.toggle('電気的安全性点検')
if toggle1:
    st.subheader('電気的安全性点検')
    st.write('**接触電流**')
    col1, col2 = st.columns(2)
    with col1:
        eli1.value = st.number_input(eli1.label, min_value=0.0, format='%.1f', step=0.1) # 接触電流（NC）
        eli1.bool = True if eli1.value <= 100 else False
    with col2:
        eli2.value = st.number_input(eli2.label, min_value=0.0, format='%.1f', step=0.1) # 接触電流（SFC）
        eli2.bool = True if eli2.value <= 500 else False
    st.write('**接地漏れ電流**')
    col1, col2 = st.columns(2)
    with col1:
        eli3.value = st.number_input(eli3.label, min_value=0.0, format='%.1f', step=0.1) # 接地漏れ電流（NC）
        eli3.bool = True if eli3.value <= 5000 else False
    with col2:
        eli4.value = st.number_input(eli4.label, min_value=0.0, format='%.1f', step=0.1) # 接地漏れ電流（SFC）
        eli4.bool = True if eli4.value <= 10000 else False
    toggle1 = st.toggle('**接地線抵抗**')
    if toggle1:
        eli5.value = st.number_input(eli5.label, min_value=0.0, format='%.3f', step=0.001) # 接地線抵抗
        eli5.bool = True if eli5.value <= 0.2 else False
st.divider()

# 外装点検
st.subheader('外装点検', help='総合評価には含まれません')
outi1.bool = st.checkbox(outi1.label) # 装置に汚れ、ひび割れ、破損がないか
outi2.bool = st.checkbox(outi2.label) # 操作パネルの傷、はく離がないか
outi3.bool = st.checkbox(outi3.label) # 部品のゆるみ、ねじ、ナットのゆるみがないか
outi4.bool = st.checkbox(outi4.label) # 電源コードの破損がないか
st.divider()

# 機能点検
st.subheader('機能点検')
fni1.bool = st.checkbox(fni1.label) # 現在時刻の確認
fni2.bool = st.checkbox(fni2.label) # 電源を入れた際にセルフチェック機能が働くか
fni3.bool = st.checkbox(fni3.label) # 動作中・異常発生時のランプが点灯・点滅するか
st.divider()

# 性能点検
st.subheader('性能点検')
st.write('**SPO2**')
col1, col2 = st.columns(2)
with col1:
    peri1.value = st.number_input(peri1.label, min_value=0, max_value=100, format='%d', step=1) # 酸素飽和度測定精度
    peri1.bool = True if 93 <= peri1.value and peri1.value <= 100 else False
with col2:
    peri2.value = st.number_input(peri2.label, min_value=0, step=1) # 脈拍数測定精度
    peri2.bool = True if 69 <= peri2.value and peri2.value <= 75 else False
st.write('**心電図**')
peri3.bool = st.checkbox(peri3.label) # 入力した波形が正常に出力されるか
st.write('**体温部**')
set1 = 'ー'
toggle_peri4 = st.toggle(peri4.label)
if toggle_peri4:
    col1, col2 = st.columns(2)
    with col1:
        set1 = st.number_input('設定値（℃）', value=37.0, format='%.1f', step=0.1)
        set1 = round(set1, 1)
        min1 = round(set1-0.1, 1)
        max1 = round(set1+0.1, 1)
    with col2:
        if '体温測定精度' not in st.session_state: st.session_state['体温測定精度'] = 0.0
        def update_peri4():
            st.session_state['体温測定精度'] = st.session_state['新体温測定精度']
        peri4.value = st.number_input(f'{peri4.label}（{str(min1)} ～ {str(max1)} ℃）', value=st.session_state['体温測定精度'], format='%.1f', step=0.1, key='新体温測定精度', on_change=update_peri4) # 体温測定精度
        st.session_state['体温測定精度'] = peri4.value
        peri4.bool = True if min1 <= peri4.value and peri4.value <= max1 else False
st.write('**血圧部**')
col1, col2 = st.columns(2)
with col1:
    set2 = st.number_input('設定値（mmHg）', value=120, step=1)
    min2 = set2-20
    max2 = set2+130
with col2:
    if '血圧部（上限）' not in st.session_state: st.session_state['血圧部（上限）'] = 0
    def update_peri5():
        st.session_state['血圧部（上限）'] = st.session_state['新血圧部（上限）']
    peri5.value = st.number_input(f'{peri5.label}（{str(min2)} ～ {str(max2)} mmHg）', value=st.session_state['血圧部（上限）'], format='%d', step=1, key='新血圧部（上限）', on_change=update_peri5) # 最高血圧警報測定(上限値)
    st.session_state['血圧部（上限）'] = peri5.value
    peri5.bool = True if min2 <= peri5.value and peri5.value <= max2 else False
col1, col2 = st.columns(2)
with col1:
    set3 = st.number_input('設定値（mmHg）', value=80, step=1)
    min3 = set3-20
    max3 = set3+70
with col2:
    if '血圧部（下限）' not in st.session_state: st.session_state['血圧部（下限）'] = 0
    def update_peri6():
        st.session_state['血圧部（下限）'] = st.session_state['新血圧部（下限）']
    peri6.value = st.number_input(f'{peri6.label}（{str(min3)} ～ {str(max3)} mmHg）', value=st.session_state['血圧部（下限）'], format='%d', step=1, key='新血圧部（下限）', on_change=update_peri6) # 最高血圧警報測定(下限値)
    st.session_state['血圧部（下限）'] = peri6.value
    peri6.bool = True if min3 <= peri6.value and peri6.value <= max3 else False
st.divider()



# 総合評価
st.subheader('総合評価')
errorList = [instance for instance in stObject.getInstances() if instance.category in eval_ctg and instance.bool == False]
# 評価
if not errorList: st.success('**合格**') 
else: 
    st.warning('**不合格**')
    st.write('不合格の項目')
    text = ['「' + error.name + '」' for error in errorList]
    st.warning(''.join(text))
st.divider()

# 備考
st.subheader('備考')
text_area1 = st.text_area(' ')
st.divider()



# ダウンロード
st.subheader('ダウンロード', help='点検報告書をExcel形式でダウンロードします')
# 初期値設定
if 'デフォルト' not in st.session_state: st.session_state['デフォルト'] = '点検報告書'
if 'ファイル名' not in st.session_state: st.session_state['ファイル名'] = '点検報告書'
if '新ファイル名' not in st.session_state: st.session_state['新ファイル名'] = '点検報告書'
if 'トグル' not in st.session_state: st.session_state['トグル'] = False

# 現在のファイル名
st.write('現在のファイル名：' + st.session_state['ファイル名'] + '.xlsx')

# トグル
def update_toggle(): # トグル変更時に
    st.session_state['ファイル名'] = st.session_state['デフォルト'] # ファイル名をデフォルトにする
toggle = st.toggle('ファイル名 変更', key='トグル', on_change=update_toggle)
if toggle:
    def update_file_name(): # 新ファイル名を入力後に（トグルが変更されるたびに起動する　バグ？）
        if st.session_state['トグル']: # トグルがONなら
            st.session_state['ファイル名'] = st.session_state['新ファイル名'] # ファイル名を新ファイル名にする
    st.text_input('新しいファイル名', key='新ファイル名', value=st.session_state['ファイル名'], on_change=update_file_name)
st.divider()

# ダウンロードボタン
# Excel 入力
def excel():
    wb = load_workbook(file_path)
    sheet = wb[sheet_name]
    
    # 合格・不合格
    bools = [instance.bool for instance in stObject.getInstances() if instance.category in ctg] # boolをすべて取得
    for i, bool in enumerate(bools, start=11):
        sheet[f'H{i}'] = ('ー' if bool == 'ー' else '合格' if bool else '不合格') # H11から入力

    # 機器情報
    sheet['E3'] = info1.value
    sheet['B5'] = info2.value
    sheet['B6'] = info3.value
    sheet['B7'] = info4.value
    sheet['B8'] = info5.value
    sheet['E8'] = info6.value
    # 電気的安全性点検
    sheet['F11'] = eli1.value
    sheet['F12'] = eli2.value
    sheet['F13'] = eli3.value
    sheet['F14'] = eli4.value
    sheet['F15'] = eli5.value
    # 性能点検
    sheet['D26'] = f'：{set1} ± 0.1℃        （'
    sheet['C27'] = f'最高血圧警報測定(上限値)：{('  '*(3-len(str(set2))))+str(str(set2))}mmHg（'
    sheet['C28'] = f'最高血圧警報測定(下限値)：{('  '*(3-len(str(set3))))+str(str(set3))}mmHg（'
    sheet['F23'] = peri1.value
    sheet['F24'] = peri2.value
    sheet['F26'] = peri4.value
    sheet['F27'] = peri5.value
    sheet['F28'] = peri6.value
    # 備考
    sheet['B29'] = text_area1
    # 総合評価
    sheet['H32'] = '合格' if not errorList else '不合格'
    # 保存
    byte_xlsx = BytesIO()
    wb.save(byte_xlsx)
    wb.close()
    byte_xlsx.seek(0)
    return byte_xlsx

file = excel()
file_name = st.session_state['ファイル名'] + '.xlsx'
mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# 初期化
if 'ダウンロードボタン' not in st.session_state: st.session_state['ダウンロードボタン'] = False
if 'プログレスバー' not in st.session_state: st.session_state['プログレスバー'] = False

# ダウンロードボタン
if st.button('作成', use_container_width=True):
    st.write(f'{file_name} を作成しました\n\n下部の「ダウンロード」からファイルをダウンロードしてください')
    st.session_state['ダウンロードボタン'] = True
if st.session_state['ダウンロードボタン']:
    download_button = st.download_button(label='ダウンロード', data=file, file_name=file_name, mime=mime, use_container_width=True)
    if download_button: st.session_state['プログレスバー'] = True

# プログレスバー
if st.session_state['プログレスバー']:
    progress_bar = st.progress(0) # 進行バーの初期化
    for i in range(100):
        progress_bar.progress(i + 1)
        time.sleep(0.01)
    st.success(file_name + ' をダウンロードしました')
    st.caption('※エラー発生時は、もう一度「ダウンロード」を押して下さい')
    st.session_state['ダウンロードボタン'] = False
    st.session_state['プログレスバー'] = False