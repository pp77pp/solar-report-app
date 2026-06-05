import streamlit as st
from pptx import Presentation
import io
import datetime
from PIL import Image

st.set_page_config(page_title="太陽能巡檢助手 Pro", layout="wide")
st.title("☀️ 太陽能維運巡檢助手 Pro (行動戰情室)")

# --- 側邊欄 ---
st.sidebar.header("📋 報告基本資訊")
site_name = st.sidebar.text_input("案場名稱", "草湖國小")
report_date = st.sidebar.date_input("巡檢日期", datetime.date.today())
main_summary = st.sidebar.text_area("封面總結敘述", "本日巡檢、清洗、螺絲檢查完畢。")

# --- 核心：動態分頁管理 ---
if 'pages' not in st.session_state:
    st.session_state.pages = [{'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None}]

def add_page():
    st.session_state.pages.append({'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None})

st.subheader("🖼️ 編輯巡檢投影片")

for i, page in enumerate(st.session_state.pages):
    with st.container(border=True):
        st.markdown(f"#### 第 {i+1} 頁內容")
        p_layout = st.radio(f"版面格式", ["單張大圖", "對比照片"], key=f"layout_{i}", horizontal=True)
        p_desc = st.text_input(f"本頁文字敘述", key=f"desc_{i}")
        
        c1, c2 = st.columns(2)
        p_img1 = c1.file_uploader(f"上傳照片 1", type=["jpg","png","jpeg","mpo"], key=f"img1_{i}")
        p_img2 = None
        if p_layout == "對比照片":
            p_img2 = c2.file_uploader(f"上傳照片 2", type=["jpg","png","jpeg","mpo"], key=f"img2_{i}")
        
        st.session_state.pages[i]['layout'] = p_layout
        st.session_state.pages[i]['desc'] = p_desc
        st.session_state.pages[i]['img1'] = p_img1
        st.session_state.pages[i]['img2'] = p_img2

if st.button("➕ 新增下一頁"):
    add_page()
    st.rerun()

# --- 格式轉換函數 (解決 MPO 報錯) ---
def process_image(uploaded_file):
    if uploaded_file is None: return None
    # 使用 PIL 打開並強制轉存為 JPEG 格式
    img = Image.open(uploaded_file)
    img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG")
    buf.seek(0)
    return buf

# --- 生成邏輯 ---
if st.button("🚀 生成報告", type="primary"):
    try:
        prs = Presentation("template.pptx")
        # 處理封面... (代碼同前)
        for i, p_data in enumerate(st.session_state.pages):
            if i + 1 >= len(prs.slides): break
            curr_slide = prs.slides[i+1]
            # 替換文字與圖片
            for shape in curr_slide.shapes:
                if shape.has_text_frame:
                    shape.text = shape.text.replace("{{SITE}}", site_name).replace("{{DESC}}", p_data['desc'])
                
                # 處理圖片 (調用 process_image)
                if shape.name == "PIC_SINGLE" and p_data['img1']:
                    img_stream = process_image(p_data['img1'])
                    curr_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
                    shape._element.getparent().remove(shape._element)
                elif shape.name == "PIC_LEFT" and p_data['img1']:
                    img_stream = process_image(p_data['img1'])
                    curr_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
                    shape._element.getparent().remove(shape._element)
                elif shape.name == "PIC_RIGHT" and p_data['img2']:
                    img_stream = process_image(p_data['img2'])
                    curr_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
                    shape._element.getparent().remove(shape._element)

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.download_button("📥 下載 PPT 報告", data=ppt_io.getvalue(), file_name=f"Report.pptx")
    except Exception as e:
        st.error(f"錯誤：{e}")
