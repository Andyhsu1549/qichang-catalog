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

st.set_page_config(page_title="淇錩科技有限公司 - 產品目錄 Demo", layout="wide")

st.title("淇錩科技有限公司｜產品目錄 Demo")
st.caption("圖片、型號、規格、材質皆由 Excel 讀取。支援搜尋、篩選、批次更新，以及一鍵輸出 PDF。")

# =====================
# 1) 基礎工具
# =====================
DEFAULT_EXCEL = "products_example.xlsx"
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

# 初始化 session_state
if "upsert_new" not in st.session_state:
    st.session_state.upsert_new = []
if "upsert_update" not in st.session_state:
    st.session_state.upsert_update = []

# =====================
# 2) 載入主檔
# =====================
if os.path.exists(DEFAULT_EXCEL):
    df = load_excel(DEFAULT_EXCEL)
else:
    df = pd.DataFrame(columns=list(REQUIRED_COLS))

# =====================
# 3) 側邊欄搜尋 & 篩選
# =====================
with st.sidebar:
    st.header("篩選條件")
    q = st.text_input("關鍵字（型號/規格/材質）")
    cats = st.multiselect("類別", sorted(df["類別"].unique().tolist())) if not df.empty else []
    mats = st.multiselect("材質", sorted(df["材質"].unique().tolist())) if not df.empty else []

    st.markdown("---")
    st.markdown("### 更多選項")
    view_mode = st.selectbox(
        "顯示模式",
        ["全部產品", "僅顯示新增的", "僅顯示更新過的"],
        index=0
    )

# 基礎篩選
filtered = df.copy()
if q:
    q_lower = q.lower()
    filtered = filtered[filtered.apply(
        lambda r: q_lower in (" ".join(r.astype(str).values)).lower(), axis=1)]
if cats:
    filtered = filtered[filtered["類別"].isin(cats)]
if mats:
    filtered = filtered[filtered["材質"].isin(mats)]

# 顯示模式 (串接 Upsert 結果)
if view_mode == "僅顯示新增的" and st.session_state.upsert_new:
    filtered = filtered[filtered["型號"].isin(st.session_state.upsert_new)]
elif view_mode == "僅顯示更新過的" and st.session_state.upsert_update:
    filtered = filtered[filtered["型號"].isin(st.session_state.upsert_update)]

st.subheader(f"產品列表（{len(filtered)} 筆）")

# =====================
# 4) 卡片式展示
# =====================
cols_per_row = 3
rows = math.ceil(len(filtered) / cols_per_row)
records = filtered.to_dict(orient="records")

for i in range(rows):
    row_cards = records[i*cols_per_row:(i+1)*cols_per_row]
    cols = st.columns(cols_per_row)
    for col, item in zip(cols, row_cards):
        with col:
            img_path = str(item.get("圖片路徑", ""))
            try:
                st.image(img_path, use_container_width=True)
            except Exception:
                st.image(Image.new("RGB",(600,400),(230,230,230)), use_container_width=True, caption="範例")
            st.markdown(f"**型號**：{item['型號']}")
            st.markdown(f"**規格**：{item['規格']}")
            st.markdown(f"**材質**：{item['材質']}")
            st.markdown(f"<span style='color:#888'>類別：{item['類別']}</span>", unsafe_allow_html=True)

st.divider()

# =====================
# 5) PDF 匯出
# =====================
def make_catalog_pdf(items):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W,H = A4
    margin = 15*mm
    def header():
        c.setFont("Helvetica-Bold",14)
        c.drawString(margin,H-margin+2*mm,"淇錩科技有限公司 產品型錄")
        c.setFont("Helvetica",9)
        c.drawString(margin,H-margin-3*mm,"（內容由 Excel 匯入，可即時更新）")
        c.line(margin,H-margin-5*mm,W-margin,H-margin-5*mm)
    header()
    y = H - margin - 12*mm
    img_max_w,img_max_h = 70*mm,45*mm
    line_gap = 6*mm
    for item in items:
        if y < margin + img_max_h + 25*mm:
            c.showPage(); header(); y = H - margin - 12*mm
        img_path = str(item.get("圖片路徑",""))
        img_reader=None
        if os.path.exists(img_path):
            try:
                img_reader=ImageReader(img_path)
                iw,ih=Image.open(img_path).size
                scale=min(img_max_w/iw,img_max_h/ih)
                dw,dh=iw*scale,ih*scale
                c.drawImage(img_reader,margin,y-dh,width=dw,height=dh,preserveAspectRatio=True,mask='auto')
            except: pass
        if not img_reader:
            c.rect(margin,y-img_max_h,img_max_w,img_max_h)
            c.setFont("Helvetica",8)
            c.drawCentredString(margin+img_max_w/2,y-img_max_h/2,"No Image")
        tx = margin+img_max_w+10*mm
        c.setFont("Helvetica-Bold",12); c.drawString(tx,y,f"型號：{item.get('型號','')}")
        c.setFont("Helvetica",11)
        c.drawString(tx,y-12,f"規格：{item.get('規格','')}")
        c.drawString(tx,y-24,f"材質：{item.get('材質','')}")
        c.setFont("Helvetica",9); c.setFillColorRGB(0.4,0.4,0.4)
        c.drawString(tx,y-36,f"類別：{item.get('類別','')}"); c.setFillColorRGB(0,0,0)
        y -= max(img_max_h,42*mm) + line_gap
    c.save(); buffer.seek(0); return buffer

st.subheader("輸出 PDF")
if st.button("產生 PDF"):
    pdf_bytes = make_catalog_pdf(filtered.to_dict(orient="records"))
    st.download_button("下載 產品型錄.pdf", data=pdf_bytes,
                       file_name="產品型錄.pdf", mime="application/pdf")

st.divider()

# =====================
# 6) 批次更新 / 新增 (Upsert)
# =====================
st.subheader("批次更新 / 新增 (Upsert)")

with st.expander("上傳更新檔 → 預覽差異 → 套用", expanded=False):
    up_file = st.file_uploader("上傳更新 Excel（需欄位：類別、型號、規格、材質、圖片路徑）", type=["xlsx"])
    if up_file:
        df_up = pd.read_excel(up_file).fillna("")
        miss = REQUIRED_COLS - set(df_up.columns)
        if miss:
            st.error("更新檔缺少欄位：" + "、".join(miss))
        else:
            # 以型號為 key
            df["_key"] = df["型號"].map(normalize_key)
            df_up["_key"] = df_up["型號"].map(normalize_key)
            key_master, key_up = set(df["_key"]), set(df_up["_key"])
            to_insert = key_up - key_master
            to_check = key_up & key_master
            updates, same = [], []
            for k in to_check:
                row_m = df.loc[df["_key"]==k, list(REQUIRED_COLS)].iloc[0]
                row_u = df_up.loc[df_up["_key"]==k, list(REQUIRED_COLS)].iloc[0]
                if any(str(row_m[c])!=str(row_u[c]) for c in REQUIRED_COLS):
                    updates.append(k)
                else:
                    same.append(k)

            st.write(f"🆕 新增：{len(to_insert)} 筆，✏️ 變更：{len(updates)} 筆，✅ 相同：{len(same)} 筆")

            # 更新 session_state，讓下拉選單能用
            st.session_state.upsert_new = df_up[df_up["_key"].isin(to_insert)]["型號"].tolist()
            st.session_state.upsert_update = df_up[df_up["_key"].isin(updates)]["型號"].tolist()

            if st.button("套用更新", type="primary"):
                if os.path.exists(DEFAULT_EXCEL):
                    bak = backup_excel(DEFAULT_EXCEL)
                    st.info(f"已自動備份：{bak}")
                base = df.set_index("_key")
                for _, row in df_up.iterrows():
                    base.loc[row["_key"], list(REQUIRED_COLS)] = row[list(REQUIRED_COLS)]
                out = base.reset_index()[list(REQUIRED_COLS)]
                save_excel(out, DEFAULT_EXCEL)
                st.success("更新完成！重新整理頁面即可查看最新清單。")
