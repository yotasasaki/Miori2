import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image
import streamlit as st


st.set_page_config(layout="wide")
st.title(':blue[History of Yota]🤠✨✨✨🤠')
#------------------写真の貼り付け(洋太)------------------
col1,col2,col3=st.columns(3)

with col1:
    st.header("いけいけ")
    image=Image.open("./data/y1.jpg")
    st.image(image,width=400)

with col2:
    st.header("withおじいちゃん")
    image=Image.open("./data/y2.jpg")
    st.image(image,width=400)

with col3:
    st.header("ジャガーズ")
    image=Image.open("./data/y3.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("たかし")
    image=Image.open("./data/y4.jpg")
    st.image(image,width=400)

with col2:
    st.header("ようた")
    image=Image.open("./data/y5.jpg")
    st.image(image,width=400)

with col3:
    st.header("なおし")
    image=Image.open("./data/y6.jpg")
    st.image(image,width=400)


col1,col2,col3=st.columns(3)

with col1:
    st.header("3兄弟")
    image=Image.open("./data/y7.jpg")
    st.image(image,width=400)

with col2:
    st.header("ないすpic")
    image=Image.open("./data/y8.jpg")
    st.image(image,width=400)

with col3:
    st.header("やんきー")
    image=Image.open("./data/y9.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("ノリノリテニス")
    image=Image.open("./data/y10.jpg")
    st.image(image,width=400)

with col2:
    st.header("卒業旅行")
    image=Image.open("./data/y11.jpg")
    st.image(image,width=400)

with col3:
    st.header("ゆうか")
    image=Image.open("./data/y12.jpg")
    st.image(image,width=400)

col1,col2,col3=st.columns(3)

with col1:
    st.header("もち")
    image=Image.open("./data/y13.jpg")
    st.image(image,width=400)

with col2:
    st.header("サル島BBQ")
    image=Image.open("./data/y14.jpg")
    st.image(image,width=400)

with col3:
    st.header("美織@軽井沢")
    image=Image.open("./data/y15.jpg")
    st.image(image,width=400)



