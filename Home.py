import streamlit as st
from PIL import Image

st.title('CheckMake')
st.write('点検報告書作成アプリ')
st.caption('Update：2024/07')
st.divider()

st.header('Info')
st.write('医療機器の点検報告書を作成し、Excel形式で保存することができます')
image = Image.open('data/点検報告書 例.png')
st.image(image, width=300)
st.divider()

st.header('Pages')
st.page_link('pages/輸液ポンプ.py', label='輸液ポンプ')