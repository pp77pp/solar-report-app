import streamlit as st
from pptx import Presentation
import io
import datetime

st.set_page_config(page_title="巡檢報告助手", layout="wide")
st.title("☀️ 太陽能維運巡檢報告助手")

# --- 側邊欄 ---
st.sidebar.header("報告基本資訊")
site_name = st.sidebar.text_input("案場名稱", "草湖國小")
report_date = st.sidebar.date_input("巡檢日期", datetime.date.today())

# --- 主畫面 ---
st.subheader("🖼️ 編輯巡檢投影片")
mode = st.radio("選擇本頁版面格式", ["單張大圖 (例如: 巡檢概況)", "兩張並排 (例如: 清洗前後對比)"])
desc_text = st.text_area("文字敘述內容", "6/4(四)於草湖國小進行作業，皆無異常。")

col1, col2 = st.columns(2)
img1 = None
img2 = None

if mode == "單張大圖 (例如: 巡檢概況)":
    img1 = st.file_uploader("請拍攝或上傳照片", type=["jpg", "png", "jpeg"], key="single")
else:
    with col1:
        img1 = st.file_uploader("上傳左側照片 (清洗前)", type=["jpg", "png", "jpeg"], key="left")
    with col2:
        img2 = st.file_uploader("上傳右側照片 (清洗後)", type=["jpg", "png", "jpeg"], key="right")

# --- 生成功能 ---
if st.button("🚀 生成報告並下載"):
    if not img1:
        st.error("請至少上傳一張照片！")
    else:
        try:
            # 讀取模板 (請確保 template.pptx 已上傳到 GitHub)
            prs = Presentation("template.pptx")
            slide = prs.slides[0] # 這裡預設修改第一張投影片

            # 1. 替換文字標籤
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if "{{SITE}}" in shape.text:
                        shape.text = shape.text.replace("{{SITE}}", site_name)
                    if "{{DESC}}" in shape.text:
                        shape.text = shape.text.replace("{{DESC}}", desc_text)

            # 2. 替換圖片函數 (根據選取範圍窗格的名稱)
            def replace_image(shape_name, uploaded_file):
                for shape in slide.shapes:
                    if shape.name == shape_name:
                        # 紀錄位置
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        # 插入新圖
                        slide.shapes.add_picture(io.BytesIO(uploaded_file.read()), left, top, width, height)
                        # 移除舊圖
                        sp = shape._element
                        sp.getparent().remove(sp)

            if mode == "單張大圖 (例如:巡檢概況)":
                replace_image("PIC_SINGLE", img1)
            else:
                replace_image("PIC_LEFT", img1)
                if img2:
                    replace_image("PIC_RIGHT", img2)

            # 下載 PPT
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            st.download_button(
                label="📥 下載生成的 PPT 報告",
                data=ppt_io.getvalue(),
                file_name=f"{site_name}_巡檢報告.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"發生錯誤：{e}。請確認 template.pptx 是否已正確上傳！")
