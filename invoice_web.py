#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書入力ツール（Web版）
Streamlit で動作するブラウザ版。
起動: streamlit run invoice_web.py
"""

import streamlit as st
from openpyxl import load_workbook
from datetime import date
from io import BytesIO
import os, re

# ─── 設定 ───
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "請求書(ひな形）.xlsx")
MAX_ITEMS   = 16
TAX_OPTIONS = ["10%", "8%", "非課税"]

# ─── ページ設定 ───
st.set_page_config(page_title="請求書入力ツール", page_icon="📄", layout="wide")

# ─── カスタムCSS ───
st.markdown("""
<style>
    /* 全体のフォント */
    .stApp { font-family: "Hiragino Sans", "Yu Gothic UI", "Meiryo", sans-serif; }

    /* ヘッダー */
    .main-header {
        background: linear-gradient(135deg, #1B5E8C 0%, #2980b9 100%);
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .main-header p  { margin: 0.3rem 0 0 0; font-size: 0.95rem; opacity: 0.85; }

    /* セクション */
    .section-box {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .section-title {
        color: #1B5E8C;
        font-size: 1.1rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
        border-bottom: 2px solid #1B5E8C;
        padding-bottom: 0.3rem;
    }

    /* 合計エリア */
    .total-box {
        background: linear-gradient(135deg, #f0f7ff 0%, #e8f4f8 100%);
        border: 2px solid #1B5E8C;
        border-radius: 10px;
        padding: 1.2rem;
        margin: 1rem 0;
    }
    .grand-total {
        font-size: 1.6rem;
        font-weight: 800;
        color: #1B5E8C;
        text-align: right;
    }

    /* データエディタのヘッダー */
    [data-testid="stDataEditor"] th {
        background: #1B5E8C !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── ヘッダー ───
st.markdown("""
<div class="main-header">
    <h1>📄 請求書入力ツール</h1>
    <p>入力して「Excelを生成」ボタンを押すだけ。ブラウザからそのままダウンロードできます。</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════
#  基本情報
# ══════════════════════════════════════════
st.markdown('<div class="section-box"><div class="section-title">📋 基本情報</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    client     = st.text_input("請求先（御中）*", placeholder="株式会社○○○")
    invoice_dt = st.date_input("請求日", value=date.today())
with col2:
    invoice_no = st.text_input("請求番号", placeholder="INV-2026-001")
    deadline   = st.text_input("支払い期限", placeholder="2026年4月30日")

subject = st.text_input("件名", placeholder="○○案件に関する請求")
st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
#  明細テーブル
# ══════════════════════════════════════════
st.markdown('<div class="section-box"><div class="section-title">📝 明細（最大16行）</div>', unsafe_allow_html=True)
st.caption("※ 税区分: 10% = 標準税率 ／ 8% = 軽減税率 ／ 非課税")

# セッション初期化
if "items" not in st.session_state:
    st.session_state.items = [
        {"税区分": "10%", "品番": "", "品名": "", "数量": None,
         "単位": "", "単価": None}
        for _ in range(MAX_ITEMS)
    ]

import pandas as pd

df = pd.DataFrame(st.session_state.items)

edited = st.data_editor(
    df,
    column_config={
        "税区分": st.column_config.SelectboxColumn(
            "税区分", options=TAX_OPTIONS, default="10%", width="small"),
        "品番":   st.column_config.TextColumn("品番", width="small"),
        "品名":   st.column_config.TextColumn("品名", width="large"),
        "数量":   st.column_config.NumberColumn("数量", min_value=0, format="%.0f", width="small"),
        "単位":   st.column_config.TextColumn("単位", width="small"),
        "単価":   st.column_config.NumberColumn("単価", min_value=0, format="%.0f", width="medium"),
    },
    use_container_width=True,
    num_rows="fixed",
    hide_index=True,
    key="item_editor",
)

# 更新を保存
st.session_state.items = edited.to_dict("records")
st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
#  合計計算
# ══════════════════════════════════════════
base10 = base8 = exempt = 0.0
for row in st.session_state.items:
    qty   = row.get("数量") or 0
    price = row.get("単価") or 0
    sub   = qty * price
    if sub == 0:
        continue
    tax = row.get("税区分", "10%")
    if tax == "8%":
        base8 += sub
    elif tax == "非課税":
        exempt += sub
    else:
        base10 += sub

subtotal = base10 + base8 + exempt
tax10    = int(base10 * 0.10)
tax8     = int(base8  * 0.08)
grand    = int(subtotal + tax10 + tax8)

st.markdown('<div class="total-box">', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    st.metric("小計", f"¥{subtotal:,.0f}")
with c2:
    st.metric("10%対象 → 消費税", f"¥{base10:,.0f} → ¥{tax10:,.0f}")
    st.metric("8%対象 → 消費税",  f"¥{base8:,.0f} → ¥{tax8:,.0f}")
with c3:
    st.metric("非課税", f"¥{exempt:,.0f}")

st.markdown(f'<div class="grand-total">御請求金額（税込）　¥{grand:,.0f}</div>',
            unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
#  備考
# ══════════════════════════════════════════
st.markdown('<div class="section-box"><div class="section-title">📝 備考・支払い条件</div>', unsafe_allow_html=True)
remarks = st.text_area("", placeholder="例: 翌月末日払い、銀行振込にてお願いいたします。",
                        height=100, label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
#  Excel 生成 ＆ ダウンロード
# ══════════════════════════════════════════
def parse_date(s):
    """YYYY/MM/DD 等をパース"""
    m = re.match(r"(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})", str(s).strip())
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    t = date.today()
    return t.year, t.month, t.day


def generate_excel() -> bytes:
    """テンプレートに書き込んでバイト列を返す"""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # 請求日
    y, mo, d = invoice_dt.year, invoice_dt.month, invoice_dt.day
    ws["H4"] = f"請求日　　{y}年{mo:02d}月{d:02d}日"

    # 請求番号
    inv = invoice_no.strip()
    ws["H5"] = f"請求番号　{inv}" if inv else "請求番号"

    # 請求先
    ws["C8"] = f"　{client.strip()}　御中"

    # 件名
    ws["D14"] = subject.strip()

    # 支払い期限
    ws["E17"] = deadline.strip()

    # 明細行 (行 22〜37)
    for i, rw in enumerate(st.session_state.items):
        r = 22 + i
        hinmei  = (rw.get("品名") or "").strip()
        qty     = rw.get("数量")
        price   = rw.get("単価")

        if not (hinmei or qty or price):
            for c in range(2, 9):
                ws.cell(row=r, column=c).value = None
            continue

        tax = rw.get("税区分", "10%")
        if tax == "8%":
            code = "8%"
        elif tax == "非課税":
            code = "非"
        else:
            code = ""

        ws.cell(row=r, column=2).value = code
        ws.cell(row=r, column=3).value = (rw.get("品番") or "").strip() or None
        ws.cell(row=r, column=4).value = hinmei or None
        ws.cell(row=r, column=5).value = float(qty) if qty else None
        ws.cell(row=r, column=6).value = (rw.get("単位") or "").strip() or None
        ws.cell(row=r, column=7).value = float(price) if price else None
        ws.cell(row=r, column=8).value = f"=E{r}*G{r}"

    # 税計算フォーミュラ
    ws["E40"] = '=SUMIF(B22:B37,"8%",H22:H37)'
    ws["G40"] = "=H38-E40-H40"

    # 備考
    if remarks.strip():
        ws["B39"] = remarks.strip()

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── ボタン ──
st.markdown("---")
col_l, col_r = st.columns([1, 1])

with col_l:
    if st.button("🗑️ フォームをクリア", use_container_width=True):
        st.session_state.items = [
            {"税区分": "10%", "品番": "", "品名": "", "数量": None,
             "単位": "", "単価": None}
            for _ in range(MAX_ITEMS)
        ]
        st.rerun()

with col_r:
    # バリデーション
    can_export = bool(client.strip()) and any(
        (r.get("品名") or "").strip() for r in st.session_state.items
    )

    if can_export:
        today_str  = date.today().strftime("%Y%m%d")
        default_nm = f"請求書_{today_str}_{client.strip()}.xlsx"
        excel_data = generate_excel()
        st.download_button(
            label="📥 Excelを生成してダウンロード",
            data=excel_data,
            file_name=default_nm,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    else:
        st.button("📥 Excelを生成してダウンロード",
                  disabled=True, use_container_width=True,
                  help="請求先と明細を1行以上入力してください")
