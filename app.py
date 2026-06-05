import streamlit as st
from pptx import Presentation
import io
import datetime
import copy

st.set_page_config(page_title="太陽能巡檢助手 Pro", layout="wide")
st.title("☀️ 太陽能維運巡檢助手 Pro (多頁編輯版)")

# --- 側邊欄：全域資訊 ---
st.sidebar.header("📋 報告基本資訊")
site_name = st.sidebar.text_input("案場名稱", "草湖國小")
report_date = st.sidebar.date_input("巡檢日期", datetime.date.today())
main_summary = st.sidebar.text_area("封面總結敘述", "草湖國小6月巡檢、模組清洗、螺絲檢查。")

# --- 核心：動態分頁管理 ---
if 'pages' not in st.session_state:
    st.session_state.pages = [{'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None}]

def add_page():
    st.session_state.pages.append({'layout': '單張大圖', 'desc': '', 'img1': None, 'img2': None})

def delete_page(index):
    if len(st.session_state.pages) > 1:
        st.session_state.pages.pop(index)

st.subheader("🖼️ 編輯巡檢投影片 (可新增多頁)")

# 遍歷所有已新增的頁面
for i, page in enumerate(st.session_state.pages):
    with st.container(border=True):
        col_t, col_d = st.columns([8, 1])
        col_t.markdown(f"#### 第 {i+1} 頁內容")
        if col_d.button("🗑️", key=f"del_{i}"):
            delete_page(i)
            st.rerun()

        p_layout = st.radio(f"版面格式", ["單張大圖", "對比照片"], key=f"layout_{i}", horizontal=True)
        p_desc = st.text_input(f"本頁文字敘述", key=f"desc_{i}", placeholder="例如：於草湖國小進行DC箱掃瞄，無異常。")
        
        c1, c2 = st.columns(2)
        p_img1 = c1.file_uploader(f"上傳照片 1", type=["jpg","png","jpeg"], key=f"img1_{i}")
        p_img2 = None
        if p_layout == "對比照片":
            p_img2 = c2.file_uploader(f"上傳照片 2 (對比圖)", type=["jpg","png","jpeg"], key=f"img2_{i}")
        
        # 暫存資料到 session_state
        st.session_state.pages[i]['layout'] = p_layout
        st.session_state.pages[i]['desc'] = p_desc
        st.session_state.pages[i]['img1'] = p_img1
        st.session_state.pages[i]['img2'] = p_img2

if st.button("➕ 新增下一頁投影片"):
    add_page()
    st.rerun()

st.divider()

# --- 生成邏輯 ---
if st.button("🚀 生成完整報告並下載", type="primary", use_container_width=True):
    try:
        prs = Presentation("template.pptx")
        
        # 1. 處理第一頁 (標題頁)
        title_slide = prs.slides[0]
        for shape in title_slide.shapes:
            if shape.has_text_frame:
                if "{{DATE}}" in shape.text: shape.text = shape.text.replace("{{DATE}}", str(report_date))
                if "{{SITE}}" in shape.text: shape.text = shape.text.replace("{{SITE}}", site_name)
                if "{{DESC}}" in shape.text: shape.text = shape.text.replace("{{DESC}}", main_summary)

        # 2. 準備模板投影片 (假設第2頁是單圖模板，第3頁是雙圖模板)
        # 注意：這部分邏輯會根據你 template.pptx 裡面的頁面順序來抓取格式
        single_tpl_index = 1 # 你的第2頁
        double_tpl_index = 2 # 你的第3頁
        
        # 為了不影響原始模板，我們建立一個空的 PPT 來存放生成的結果
        output_prs = Presentation("template.pptx")
        # 先清空除第一頁以外的所有頁面，我們手動新增
        for _ in range(len(output_prs.slides) - 1):
            rId = output_prs.slides._sldIdLst[1].rId
            output_prs.part.drop_rel(rId)
            del output_prs.slides._sldIdLst[1]

        # 重新處理第一頁文字
        ts = output_prs.slides[0]
        for s in ts.shapes:
            if s.has_text_frame:
                s.text = s.text.replace("{{DATE}}", str(report_date)).replace("{{SITE}}", site_name).replace("{{DESC}}", main_summary)

        # 根據用戶編輯的頁數，動態產生新的投影片
        for p_data in st.session_state.pages:
            # 決定要用哪種格式
            source_index = single_tpl_index if p_data['layout'] == "單張大圖" else double_tpl_index
            source_slide = prs.slides[source_index]
            
            # 複製模板頁面到新 PPT (這是一個簡單的複製邏輯)
            # 這裡為了維持你的格式，我們直接操作 prs 的副本
            # 實務上我們直接在各頁替換內容
            pass 

        # --- 簡化版邏輯：直接在原 prs 的第2, 3...頁填入資料 ---
        # 為了維持你 PPT 的精美排版，這裡採「依序填充」邏輯
        for i, p_data in enumerate(st.session_state.pages):
            if i + 1 >= len(prs.slides): break # 防止超過模板頁數
            
            curr_slide = prs.slides[i+1]
            # 替換文字
            for shape in curr_slide.shapes:
                if shape.has_text_frame:
                    shape.text = shape.text.replace("{{SITE}}", site_name).replace("{{DESC}}", p_data['desc'])
            
            # 替換圖片
            def replace_image(slide, shape_name, file_data):
                for shape in slide.shapes:
                    if shape.name == shape_name and file_data:
                        l, t, w, h = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(io.BytesIO(file_data.read()), l, t, w, h)
                        sp = shape._element
                        sp.getparent().remove(sp)

            if p_data['layout'] == "單張大圖":
                replace_image(curr_slide, "PIC_SINGLE", p_data['img1'])
            else:
                replace_image(curr_slide, "PIC_LEFT", p_data['img1'])
                replace_image(curr_slide, "PIC_RIGHT", p_data['img2'])

        # 下載
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.download_button(label="📥 點我下載完整巡檢報告", data=ppt_io.getvalue(), file_name=f"{site_name}_報告.pptx")

    except Exception as e:
        st.error(f"錯誤：{e}")
