import streamlit as st
from pptx import Presentation
import io
import datetime
from PIL import Image

# 網頁環境設定
st.set_page_config(page_title="太陽能巡檢戰情室", layout="wide")
st.title("☀️ 太陽能巡檢報告產出 (格式完全保留版)")

# --- 側邊欄 ---
st.sidebar.header("📋 報告基本資訊")
site_name = st.sidebar.text_input("案場名稱", "草湖國小")
report_date = st.sidebar.date_input("巡檢日期", datetime.date.today())
main_summary = st.sidebar.text_area("首頁工作摘要 ({{DESC}})", "草湖國小6月巡檢、模組清洗、螺絲檢查。")

# --- 頁面管理 ---
if 'pages' not in st.session_state:
    st.session_state.pages = [{'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None}]

st.subheader("🖼️ 內容頁編輯 (自動從第2頁開始填入)")

for i, page in enumerate(st.session_state.pages):
    with st.container(border=True):
        st.markdown(f"#### 第 {i+1} 份巡檢內容")
        p_layout = st.radio(f"格式", ["單張大圖", "兩張並排"], key=f"lo_{i}", horizontal=True)
        p_desc = st.text_area(f"下方大方格敘述 ({{{{DESC}} }})", key=f"ds_{i}", height=100)
        
        c1, c2 = st.columns(2)
        p_img1 = c1.file_uploader(f"照片 1", type=["jpg","png","jpeg","mpo"], key=f"i1_{i}")
        p_img2 = None
        if p_layout == "兩張並排":
            p_img2 = c2.file_uploader(f"照片 2", type=["jpg","png","jpeg","mpo"], key=f"i2_{i}")
        
        st.session_state.pages[i].update({'layout': p_layout, 'desc': p_desc, 'img1': p_img1, 'img2': p_img2})

if st.button("➕ 新增下一頁"):
    st.session_state.pages.append({'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None})
    st.rerun()

# --- 核心：格式保留文字替換函數 ---
def robust_replace(slide, target, replacement):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                # 檢查整段文字是否包含標籤
                if target in paragraph.text:
                    # 邏輯：保留第一個 Run 的格式，將替換後的文字塞入，並清空其餘 Run
                    # 這能確保原本的顏色、大小、字體被繼承
                    combined_text = "".join(run.text for run in paragraph.runs)
                    new_text = combined_text.replace(target, str(replacement))
                    
                    if paragraph.runs:
                        # 記住原本的格式屬性
                        first_run = paragraph.runs[0]
                        # 執行替換
                        first_run.text = new_text
                        # 移除該段落其餘的 Run (避免重複顯示)
                        for r in paragraph.runs[1:]:
                            r.text = ""

# --- 圖片處理 ---
def fix_img(uploaded_file):
    if not uploaded_file: return None
    img = Image.open(uploaded_file)
    img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG")
    buf.seek(0)
    return buf

# --- 生成邏輯 ---
if st.button("🚀 生成報告並下載", type="primary", use_container_width=True):
    try:
        prs = Presentation("template.pptx")
        
        # 1. 處理封面 (第1頁)
        title_slide = prs.slides[0]
        robust_replace(title_slide, "{{SITE}}", site_name)
        robust_replace(title_slide, "{{DATE}}", str(report_date))
        robust_replace(title_slide, "{{DESC}}", main_summary)

        # 2. 處理內容頁 (從第2頁開始)
        for i, p_data in enumerate(st.session_state.pages):
            target_idx = i + 1
            if target_idx >= len(prs.slides): break
            slide = prs.slides[target_idx]
            
            # 替換文字標籤 (繼承原本 PPT 的字體顏色大小)
            robust_replace(slide, "{{SITE}}", site_name)
            robust_replace(slide, "{{DATE}}", str(report_date))
            robust_replace(slide, "{{DESC}}", p_data['desc'])

            # 替換圖片
            def replace_pic(slide, shape_name, file_data):
                for shape in slide.shapes:
                    if shape.name == shape_name and file_data:
                        l, t, w, h = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(fix_img(file_data), l, t, w, h)
                        shape._element.getparent().remove(shape._element)
                        break

            if p_data['layout'] == "單張大圖":
                replace_pic(slide, "PIC_SINGLE", p_data['img1'])
            else:
                replace_pic(slide, "PIC_LEFT", p_data['img1'])
                replace_pic(slide, "PIC_RIGHT", p_data['img2'])

        # 下載
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.download_button("📥 點我下載正式巡檢報告", data=ppt_io.getvalue(), file_name=f"{site_name}_報告.pptx")
        st.balloons()
    except Exception as e:
        st.error(f"生成失敗：{e}")
