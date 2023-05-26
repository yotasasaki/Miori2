import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image
import streamlit as st


st.set_page_config(layout="wide")
st.title(':blue[History of Miori]😇✨✨✨😇')
st.write('作成日:2023/03/25')


#------------------写真の貼り付け(美織)------------------
col1,col2,col3=st.columns(3)

with col1:
    st.header("ぐーすか")
    image=Image.open("./data/1.jpg")
    st.image(image,width=400)

with col2:
    st.header("かわゆす")
    image=Image.open("./data/2.jpg")
    st.image(image,width=400)

with col3:
    st.header("初チャリ")
    image=Image.open("./data/3.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("タワー!")
    image=Image.open("./data/4.jpg")
    st.image(image,width=400)

with col2:
    st.header("コアラ風間")
    image=Image.open("./data/5.jpg")
    st.image(image,width=400)

with col3:
    st.header("草風間")
    image=Image.open("./data/6.jpg")
    st.image(image,width=400)


col1,col2,col3=st.columns(3)

with col1:
    st.header("キツネ風間")
    image=Image.open("./data/7.jpg")
    st.image(image,width=400)

with col2:
    st.header("わあ")
    image=Image.open("./data/8.jpg")
    st.image(image,width=400)

with col3:
    st.header("殺人現場")
    image=Image.open("./data/9.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("たまにはやり返すぜい")
    image=Image.open("./data/10.jpg")
    st.image(image,width=400)

with col2:
    st.header("イケイケ兄貴")
    image=Image.open("./data/11.jpg")
    st.image(image,width=400)

with col3:
    st.header("満点")
    image=Image.open("./data/12.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("満点2")
    image=Image.open("./data/13.jpg")
    st.image(image,width=400)

with col2:
    st.header("中学")
    image=Image.open("./data/14.jpg")
    st.image(image,width=400)

with col3:
    st.header("かそぱ")
    image=Image.open("./data/15.jpg")
    st.image(image,width=400)



col1,col2,col3=st.columns(3)

with col1:
    st.header("イケイケ")
    image=Image.open("./data/16.jpg")
    st.image(image,width=400)

with col2:
    st.header("イケイケ2")
    image=Image.open("./data/17.jpg")
    st.image(image,width=400)

with col3:
    st.header("入籍記念旅行")
    image=Image.open("./data/18.jpg")
    st.image(image,width=400)














