# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import mm
import os, math, datetime as dt

# =====================
# 基本設定 & 多語字典
# =====================
st.set_page_config(page_title="Product Catalog Demo", layout="wide")

TEXT = {
    "title": {"中文": "淇錩科技有限公司｜產品目錄 Demo", "English": "Qichang Technology｜Product Catalog Demo"},
    "caption": {"中文": "圖片、型號、規格、材質皆由 Excel 讀取。支援搜尋、篩選、批次更新、PDF 匯出與中英文切換。",
                "English": "Images, model, specs and material are loaded from Excel. Supports search, filter, batch upsert, PDF export, and bilingual UI."},
    "sidebar_filters": {"中文": "篩選條件", "English": "Filters"},
    "search": {"中文": "關鍵字（型號/規格/材質）", "English": "Search (Model / Spec / Material)"},
    "category": {"中文": "類別", "English": "Category"},
    "material": {"中文": "材質", "English": "Material"},
    "more_options": {"中文": "更多選項", "English": "More Options"},
    "view_mode": {"中文": "顯示模式", "English": "View Mode"},
    "view_all": {"中文": "全部產品", "English": "All Products"},
    "view_new": {"中文": "僅顯示新增的", "English": "Only New"},
    "view_updated": {"中文": "僅顯示更新過的", "English": "Only Updated"},
    "list_header": {"中文": "產品列表（{n} 筆）", "English": "Product List ({n} items)"},
    "model": {"中文": "型號", "English": "Model"},
    "spec": {"中文": "規格", "English": "Spec"},
    "mat": {"中文": "材質", "English": "Material"},
    "cat": {"中文": "類別", "English": "Category"},
    "img_failed": {"中文": "測試", "English": "TEST"},
    "export_pdf": {"中文": "輸出 PDF", "English": "Export PDF"},
    "export_desc": {"中文": "將目前篩選後的清單輸出為產品型錄 PDF。",
                    "English": "Export the filtered list as a catalog PDF."},
    "generate_pdf": {"中文": "產生 PDF", "English": "Generate PDF"},
    "download_pdf": {"中文": "下載 產品型錄.pdf", "English": "Download Catalog.pdf"},
    "pdf_header_main": {"中文": "淇錩科技有限公司 產品型錄", "English": "Qichang Technology Product Catalog"},
    "pdf_header_sub": {"中文": "（內容由 Excel 匯入，可即時更新）", "English": "(Content imported from Excel, updates in real time)"},
    "pdf_no_image": {"中文": "無圖片", "English": "No Image"},
    "upsert_section": {"中文": "批次更新 / 新增 (Upsert)", "English": "Batch Update / Insert (Upsert)"},
    "upsert_expander": {"中文": "上傳更新檔 → 預覽差異 → 套用", "English": "Upload Update File → Preview Diff → Apply"},
    "upsert_uploader": {"中文": "上傳更新 Excel（需欄位：類別、型號、規格、材質、圖片路徑）",
                        "English": "Upload update Excel (columns required: Category, Model, Spec, Material, ImagePath)"},
    "missing_cols": {"中文": "更新檔缺少欄位：", "English": "Missing columns in update file: "},
    "diff_counts": {"中文": "🆕 新增：{a} 筆，✏️ 變更：{b} 筆，✅ 相同：{c} 筆",
                    "English": "🆕 New: {a}  | ✏️ Updated: {b}  | ✅ Unchanged: {c}"},
    "apply_update": {"中文": "套用更新", "English": "Apply Update"},
    "backup_info": {"中文": "已自動備份：", "English": "Backup created: "},
    "update_done": {"中文": "更新完成！重新整理頁面即可查看最新清單。", "English": "Update completed! Refresh to see the latest list."},
    "excel_not_found": {"中文": "找不到 {f}，請先放置於專案根目錄。",
                        "English": "Cannot find {f}. Please place it in the project root."}
}

def T(key, lang): return TEXT[key][lang]

# 語言選擇（放在 sidebar 最上方）
lang = st.sidebar.selectbox("語言 / Language", ["中文", "English"], index=0)

st.title(T("title", lang))
st.caption(T("caption", lang))

# =====================
# 基礎工具
# =====================
DEFAULT_EXCEL = "products_example.xlsx"  # 預設讀取這份
REQUIRED_COLS = {"類別","型號","規格","材質","圖片路徑"}

def load_excel(path):
    return pd.read_excel(path).fillna("")

def save_excel(df, path):
    df.to_excel(path, index=False)

def backup_excel(path):
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.replace(".xlsx", f"_{ts}.xlsx")
    os.replace(path, bak)
    return bak

def normalize_key(s):
    return str(s).strip().lower()

# =====================
# 載入主檔
# =====================
if os.path.exists(DEFAULT_EXCEL):
    df = load_excel(DEFAULT_EXCEL)
else:
    st.error(T("excel_not_found", lang).format(f=DEFAULT_EXCEL))
    df = pd.DataFrame(columns=list(REQUIRED_COLS))

# 初始化 session_state（記錄本次 Upsert 的新增/更新清單）
if "upsert_new" not in st.session_state:
    st.session_state.upsert_new = []
if "upsert_update" not in st.session_state:
    st.session_state.upsert_update = []

# =====================
# 側邊欄搜尋 & 篩選 & 顯示模式
# =====================
with st.sidebar:
    st.header(T("sidebar_filters", lang))
    q = st.text_input(T("search", lang))
    cats = st.multiselect(T("category", lang),
                          sorted(df["類別"].unique().tolist())) if not df.empty else []
    mats = st.multiselect(T("material", lang),
                          sorted(df["材質"].unique().tolist())) if not df.empty else []

    st.markdown("---")
    st.markdown(f"### {T('more_options', lang)}")
    view_options = [T("view_all", lang), T("view_new", lang), T("view_updated", lang)]
    view_mode = st.selectbox(T("view_mode", lang), view_options, index=0)

# 基礎篩選
filtered = df.copy()
if not df.empty and q:
    q_lower = q.lower()
    filtered = filtered[filtered.apply(
        lambda r: q_lower in (" ".join(r.astype(str).values)).lower(), axis=1)]
if cats:
    filtered = filtered[filtered["類別"].isin(cats)]
if mats:
    filtered = filtered[filtered["材質"].isin(mats)]

# 顯示模式（串接 Upsert 結果）
if view_mode == T("view_new", lang) and st.session_state.upsert_new:
    filtered = filtered[filtered["型號"].isin(st.session_state.upsert_new)]
elif view_mode == T("view_updated", lang) and st.session_state.upsert_update:
    filtered = filtered[filtered["型號"].isin(st.session_state.upsert_update)]

st.subheader(T("list_header", lang).format(n=len(filtered)))

# =====================
# 卡片式展示
# =====================
cols_per_row = 3
rows = math.ceil(len(filtered) / cols_per_row) if len(filtered) else 0
records = filtered.to_dict(orient="records") if len(filtered) else []

for i in range(rows):
    row_cards = records[i*cols_per_row:(i+1)*cols_per_row]
    cols = st.columns(cols_per_row)
    for col, item in zip(cols, row_cards):
        with col:
            img_path = str(item.get("圖片路徑", ""))
            try:
                st.image(img_path, use_container_width=True)
            except Exception:
                st.image(Image.new("RGB",(600,400),(230,230,230)),
                         use_container_width=True, caption=T("img_failed", lang))
            st.markdown(f"**{T('model', lang)}**：{item['型號']}")
            st.markdown(f"**{T('spec', lang)}**：{item['規格']}")
            st.markdown(f"**{T('mat', lang)}**：{item['材質']}")
            st.markdown(f"<span style='color:#888'>{T('cat', lang)}：{item['類別']}</span>",
                        unsafe_allow_html=True)

st.divider()

# =====================
# PDF 匯出（多語）
# =====================
def make_catalog_pdf(items, lang_sel):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4
    margin = 15 * mm

    def header():
        c.setFont("Helvetica-Bold", 14)
        c.drawString(margin, H - margin + 2*mm, T("pdf_header_main", lang_sel))
        c.setFont("Helvetica", 9)
        c.drawString(margin, H - margin - 3*mm, T("pdf_header_sub", lang_sel))
        c.line(margin, H - margin - 5*mm, W - margin, H - margin - 5*mm)

    header()
    y = H - margin - 12*mm

    img_max_w = 70 * mm
    img_max_h = 45 * mm
    line_gap = 6 * mm

    for item in items:
        if y < margin + img_max_h + 25*mm:
            c.showPage(); header(); y = H - margin - 12*mm

        img_path = str(item.get("圖片路徑",""))
        img_reader = None
        if os.path.exists(img_path):
            try:
                img_reader = ImageReader(img_path)
                iw, ih = Image.open(img_path).size
                scale = min(img_max_w/iw, img_max_h/ih)
                dw, dh = iw*scale, ih*scale
                c.drawImage(img_reader, margin, y - dh, width=dw, height=dh,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                img_reader = None

        if not img_reader:
            c.rect(margin, y - img_max_h, img_max_w, img_max_h)
            c.setFont("Helvetica", 8)
            c.drawCentredString(margin + img_max_w/2, y - img_max_h/2, T("pdf_no_image", lang_sel))

        tx = margin + img_max_w + 10*mm
        c.setFont("Helvetica-Bold", 12)
        if lang_sel == "中文":
            c.drawString(tx, y, f"{T('model', lang_sel)}：{item.get('型號','')}")
            c.setFont("Helvetica", 11)
            c.drawString(tx, y - 12, f"{T('spec', lang_sel)}：{item.get('規格','')}")
            c.drawString(tx, y - 24, f"{T('mat', lang_sel)}：{item.get('材質','')}")
            c.setFont("Helvetica", 9); c.setFillColorRGB(0.4,0.4,0.4)
            c.drawString(tx, y - 36, f"{T('cat', lang_sel)}：{item.get('類別','')}")
            c.setFillColorRGB(0,0,0)
        else:
            c.drawString(tx, y, f"{T('model', lang_sel)}: {item.get('型號','')}")
            c.setFont("Helvetica", 11)
            c.drawString(tx, y - 12, f"{T('spec', lang_sel)}: {item.get('規格','')}")
            c.drawString(tx, y - 24, f"{T('mat', lang_sel)}: {item.get('材質','')}")
            c.setFont("Helvetica", 9); c.setFillColorRGB(0.4,0.4,0.4)
            c.drawString(tx, y - 36, f"{T('cat', lang_sel)}: {item.get('類別','')}")
            c.setFillColorRGB(0,0,0)

        y -= max(img_max_h, 42*mm) + line_gap

    c.save()
    buffer.seek(0)
    return buffer

st.subheader(T("export_pdf", lang))
st.write(T("export_desc", lang))
if st.button(T("generate_pdf", lang)):
    pdf_bytes = make_catalog_pdf(filtered.to_dict(orient="records"), lang)
    # 依語言決定檔名
    fname = "產品型錄.pdf" if lang == "中文" else "Catalog.pdf"
    st.download_button(T("download_pdf", lang), data=pdf_bytes, file_name=fname, mime="application/pdf")

st.divider()

# =====================
# 批次更新 / 新增 (Upsert)
# =====================
st.subheader(T("upsert_section", lang))

with st.expander(T("upsert_expander", lang), expanded=False):
    up_file = st.file_uploader(T("upsert_uploader", lang), type=["xlsx"])
    if up_file:
        df_up = pd.read_excel(up_file).fillna("")
        miss = REQUIRED_COLS - set(df_up.columns)
        if miss:
            st.error(T("missing_cols", lang) + "、".join(miss))
        else:
            # 以型號為 key
            if not df.empty:
                df["_key"] = df["型號"].map(normalize_key)
            else:
                df["_key"] = []
            df_up["_key"] = df_up["型號"].map(normalize_key)

            key_master = set(df["_key"].tolist()) if len(df) else set()
            key_up = set(df_up["_key"].tolist())

            to_insert = key_up - key_master
            to_check = key_up & key_master

            updates, same = [], []
            for k in to_check:
                row_m = df.loc[df["_key"]==k, list(REQUIRED_COLS)].iloc[0]
                row_u = df_up.loc[df_up["_key"]==k, list(REQUIRED_COLS)].iloc[0]
                if any(str(row_m[c]) != str(row_u[c]) for c in REQUIRED_COLS):
                    updates.append(k)
                else:
                    same.append(k)

            # 顯示統計
            st.write(T("diff_counts", lang).format(a=len(to_insert), b=len(updates), c=len(same)))

            # 更新 session_state，供「顯示模式」使用
            st.session_state.upsert_new = df_up[df_up["_key"].isin(to_insert)]["型號"].tolist()
            st.session_state.upsert_update = df_up[df_up["_key"].isin(updates)]["型號"].tolist()

            if st.button(T("apply_update", lang), type="primary"):
                # 備份（若主檔存在）
                if os.path.exists(DEFAULT_EXCEL):
                    bak = backup_excel(DEFAULT_EXCEL)
                    st.info(T("backup_info", lang) + bak)

                # 合併：先以目前 df 為基礎
                if len(df):
                    base = df.set_index("_key")
                else:
                    # 如果原本沒有資料，直接以更新檔為主
                    base = pd.DataFrame(columns=list(REQUIRED_COLS) + ["_key"]).set_index("_key")

                # Upsert：有就覆蓋，沒有就新增
                for _, row in df_up.iterrows():
                    base.loc[row["_key"], list(REQUIRED_COLS)] = row[list(REQUIRED_COLS)]

                out = base.reset_index()[list(REQUIRED_COLS)].fillna("")
                out.sort_values(by=["類別", "型號"], inplace=True)
                save_excel(out, DEFAULT_EXCEL)
                st.success(T("update_done", lang))
