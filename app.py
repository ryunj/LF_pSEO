# -*- coding: utf-8 -*-
"""Streamlit 래퍼 — dashboard.html 을 화면 가득 띄운다.

실적 데이터는 저장소에 없다. data.js 가 있으면 함께 주입하고,
없으면 대시보드의 빈 화면에서 CSV 를 끌어다 놓아 쓴다.
"""
import os
import streamlit as st
import streamlit.components.v1 as components

HERE = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(page_title="pSEO 실적 대시보드", layout="wide", initial_sidebar_state="collapsed")
st.markdown(
    "<style>[data-testid='stAppViewContainer']>.main{padding:0}"
    "[data-testid='stHeader'],footer{display:none}"
    ".block-container{padding:0;max-width:100%}</style>",
    unsafe_allow_html=True,
)


def read(name):
    p = os.path.join(HERE, name)
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


html = read("dashboard.html")
data = read("data.js")
if data:
    html = html.replace('<script src="data.js"></script>', "<script>%s</script>" % data)

components.html(html, height=1400, scrolling=True)
