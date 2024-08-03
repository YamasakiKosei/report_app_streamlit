import streamlit as st
from PIL import Image

st.title('Check Make')
st.write('点検報告書作成アプリ')
st.caption('Update：2024/08')
st.divider()

st.header('Info')
st.write('医療機器の点検報告書を作成し、Excel形式で保存することができます')
image = Image.open('data/点検報告書 例.png')
st.image(image)
st.divider()

st.header('Pages')
st.page_link('pages/輸液ポンプ.py', label='輸液ポンプ')
st.page_link('pages/シリンジポンプ.py', label='シリンジポンプ')
st.page_link('pages/ベッドサイドモニター.py', label='ベッドサイドモニター')
st.page_link('pages/電気メス.py', label='電気メス')
st.page_link('pages/心電図.py', label='心電図')
st.page_link('pages/除細動器.py', label='除細動器')
st.page_link('pages/超音波手術装置.py', label='超音波手術装置')