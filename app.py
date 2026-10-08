import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import csv
import re
import io
import os
from datetime import datetime

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="글로벌 원소재 시세 Trend 대시보드",
    page_icon="📊",
    layout="wide"
)

DB_PATH = "raw_materials_trend.db"

# ---------------------------------------------------------
# 1. 5대 표준 원소재 화이트리스트 정의
# ---------------------------------------------------------
STANDARDIZED_MATERIALS = [
    "Cu",
    "ALDC 12",
    "PP-TD20 (JF1512)",
    "PP-TD+GF (JF1514G10)",
    "PP-GF30% (JF5513)",
    "PP-GF20 (G-152)",
    "PP-GF30 (G-153B)",
    "PP (JI350N)",
    "PP (JI350G)",
    "PA66-GF30 (KN333G30BL)",
    "PBT-GF30 (KP213G30BL)",
    "PP-GF20 (LGP2200-NP)",
    "PP-GF20 (LGP2200-BK)",
    "PP-GF10 (LGP2100-NP)",
    "황동 C1100",
    "Sn"
]

def format_price_label(price, curr):
    if pd.isna(price) or price is None:
        return "-"
    symbol = "₩" if curr == "KRW" else ("₹" if curr == "INR" else ("¥" if curr == "CNY" else ("$" if curr == "USD" else "")))
    if curr == "KRW":
        return f"{symbol}{price:,.0f}" if price == int(price) else f"{symbol}{price:,.2f}"
    else:
        return f"{symbol}{price:,.2f}"

# ---------------------------------------------------------
# 2. SQLite DB 초기화 & 저장 (기존 구버전 스키마 자동 초기화)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='material_prices'")
    if cur.fetchone():
        cur.execute("PRAGMA table_info(material_prices)")
        cols = [c[1] for c in cur.fetchall()]
        if 'material_name' not in cols or 'std_category' not in cols:
            cur.execute("DROP TABLE material_prices")
            
    cur.execute('''
        CREATE TABLE IF NOT EXISTS material_prices (
            subsidiary TEXT,
            std_category TEXT,
            material_name TEXT,
            month TEXT,
            price REAL,
            currency TEXT,
            supplier TEXT,
            updated_at TEXT,
            PRIMARY KEY (subsidiary, material_name, month)
        )
    ''')
    conn.commit()
    conn.close()

def save_records_to_db(df_records):
    if df_records.empty:
        return 0
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    count = 0
    for _, row in df_records.iterrows():
        if row['material_name'] not in STANDARDIZED_MATERIALS:
            continue
        try:
            p_val = float(row['단가']) if pd.notna(row['단가']) else None
        except Exception:
            p_val = None
            
        cur.execute('''
            INSERT INTO material_prices (
                subsidiary, std_category, material_name, month, price, currency, supplier, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(subsidiary, material_name, month) DO UPDATE SET
                std_category=excluded.std_category,
                price=excluded.price,
                currency=excluded.currency,
                supplier=excluded.supplier,
                updated_at=excluded.updated_at
        ''', (
            str(row['법인']), str(row['std_category']), str(row['material_name']),
            str(row['월']), p_val, str(row['통화']),
            str(row.get('공급사', '')), now_str
        ))
        count += 1
    conn.commit()
    conn.close()
    return count

def load_data_from_db():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query("SELECT * FROM material_prices ORDER BY std_category, material_name, month", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    return df

# ---------------------------------------------------------
# 3. 엑셀 원본 파일 파서 (한국/인도 엑셀 자동 매핑)
# ---------------------------------------------------------
def parse_raw_material_excel(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    records = []

    if '요약_Summary' in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name='요약_Summary', header=None)
        
        exact_mapping = {
            6: ("1. Cu", "Cu"),
            9: ("2. ALDC 12", "ALDC 12"),
            10: ("3. 사출 원재료", "PP-TD20 (JF1512)"),
            11: ("3. 사출 원재료", "PP-TD+GF (JF1514G10)"),
            12: ("3. 사출 원재료", "PP-GF30% (JF5513)"),
            13: ("3. 사출 원재료", "PP-GF20 (G-152)"),
            14: ("3. 사출 원재료", "PP-GF30 (G-153B)"),
            15: ("3. 사출 원재료", "PP (JI350N)"),
            16: ("3. 사출 원재료", "PP (JI350G)"),
            17: ("3. 사출 원재료", "PA66-GF30 (KN333G30BL)"),
            18: ("3. 사출 원재료", "PBT-GF30 (KP213G30BL)"),
            19: ("3. 사출 원재료", "PP-GF20 (LGP2200-NP)"),
            20: ("3. 사출 원재료", "PP-GF20 (LGP2200-BK)"),
            21: ("3. 사출 원재료", "PP-GF10 (LGP2100-NP)"),
            22: ("4. 황동 원소재", "황동 C1100"),
            23: ("5. Sn", "Sn")
        }
        
        subs_cols = {
            '한국': {'7월': 4, '8월': 8, '9월': 12, 'curr': 'KRW'},
            '중국': {'7월': 16, '8월': 20, '9월': 24, 'curr': 'CNY'},
            '인도': {'7월': 28, '8월': 32, '9월': 36, 'curr': 'INR'}
        }
        
        for r, (cat, mat) in exact_mapping.items():
            if r >= len(df):
                continue
            for sub, m_info in subs_cols.items():
                for m_name in ['7월', '8월', '9월']:
                    col_p = m_info[m_name]
                    price = df.iloc[r, col_p]
                    curr = df.iloc[r, col_p + 1]
                    supp = df.iloc[r, col_p + 2]
                    
                    try:
                        p_val = float(price) if pd.notna(price) else None
                    except Exception:
                        p_val = None
                        
                    supp_str = str(supp).strip() if pd.notna(supp) else ""
                    if "M/s." in supp_str:
                        m = re.search(r'M/s\.\s*([^-\n\r]+)', supp_str)
                        if m:
                            supp_str = m.group(1).strip()
                    elif "RM Supplier" in supp_str:
                        supp_str = supp_str.split('\n')[0].replace('RM Supplier -', '').strip()
                        
                    records.append({
                        '법인': sub,
                        'std_category': cat,
                        'material_name': mat,
                        '월': m_name,
                        '단가': p_val,
                        '통화': str(curr).strip() if pd.notna(curr) else m_info['curr'],
                        '공급사': supp_str
                    })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 4. 초강력 복사/붙여넣기 파서 (인도 루피 기호/줄바꿈 완벽 지원)
# ---------------------------------------------------------
def parse_smart_excel_clipboard(pasted_text, sub_default="한국"):
    if not pasted_text or not pasted_text.strip():
        return pd.DataFrame()

    sep = '\t' if '\t' in pasted_text else ','
    try:
        reader = csv.reader(io.StringIO(pasted_text), delimiter=sep)
        rows = [r for r in reader if any(c.strip() for c in r)]
    except Exception:
        rows = [[c.strip() for c in l.split(sep)] for l in pasted_text.strip().splitlines() if l.strip()]

    default_months = ["7월", "8월", "9월", "10월", "11월", "12월"]
    records = []

    for row in rows:
        if not any(row):
            continue

        # 모든 셀에서 통화 기호(₹, ₩, $, ¥) 및 쉼표 제거 후 숫자 추출
        price_cells = []
        for i, cell in enumerate(row):
            clean = re.sub(r'[^\d.]', '', str(cell))
            if clean and clean != '.':
                try:
                    val = float(clean)
                    if val > 0:
                        price_cells.append((i, val, str(cell)))
                except:
                    pass

        if not price_cells:
            continue

        # 첫 번째 숫자 앞에 있는 셀들을 원소재 이름으로 인식
        first_num_idx = price_cells[0][0]
        name_tokens = [c.split('/')[0].strip() for c in row[:first_num_idx] if c.strip()]
        name_str = " ".join(name_tokens).upper()

        std_cat, std_mat = None, None
        if "CU" in name_str or "구리" in name_str or "동" in name_str: std_cat, std_mat = "1. Cu", "Cu"
        elif "ALDC" in name_str or "알루미늄" in name_str: std_cat, std_mat = "2. ALDC 12", "ALDC 12"
        elif "JF1512" in name_str or "PP-TD20" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-TD20 (JF1512)"
        elif "JF1514" in name_str or "TD+GF" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-TD+GF (JF1514G10)"
        elif "JF5513" in name_str or "GF30%" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF30% (JF5513)"
        elif "G-152" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF20 (G-152)"
        elif "G-153B" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF30 (G-153B)"
        elif "JI350N" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP (JI350N)"
        elif "JI350G" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP (JI350G)"
        elif "KN333" in name_str or "PA66-GF30" in name_str: std_cat, std_mat = "3. 사출 원재료", "PA66-GF30 (KN333G30BL)"
        elif "KP213" in name_str or "PBT-GF30" in name_str: std_cat, std_mat = "3. 사출 원재료", "PBT-GF30 (KP213G30BL)"
        elif "LGP2200-BK" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF20 (LGP2200-BK)"
        elif "LGP2200-NP" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF20 (LGP2200-NP)"
        elif "LGP2100" in name_str: std_cat, std_mat = "3. 사출 원재료", "PP-GF10 (LGP2100-NP)"
        elif "C1100" in name_str or "황동" in name_str or "BRASS" in name_str: std_cat, std_mat = "4. 황동 원소재", "황동 C1100"
        elif "SN" in name_str or "주석" in name_str: std_cat, std_mat = "5. Sn", "Sn"

        if not std_mat:
            continue

        for m_idx, (col_idx, p_val, raw_cell) in enumerate(price_cells):
            curr = "INR" if sub_default == "인도" else ("KRW" if sub_default == "한국" else "CNY")
            if "₹" in raw_cell or "INR" in raw_cell: curr = "INR"
            elif "₩" in raw_cell or "KRW" in raw_cell: curr = "KRW"
            elif "¥" in raw_cell or "CNY" in raw_cell: curr = "CNY"
            elif "$" in raw_cell or "USD" in raw_cell: curr = "USD"
            elif col_idx + 1 < len(row) and row[col_idx + 1] in ['KRW', 'INR', 'USD', 'CNY']:
                curr = row[col_idx + 1]

            supp = ""
            for check_idx in [col_idx + 1, col_idx + 2, col_idx + 3]:
                if check_idx < len(row):
                    txt = row[check_idx]
                    if any(kw in txt for kw in ['M/s.', 'Hindalco', 'Supplier', '평화', '상농', '롯데', 'BGF', '원영', '한국특산', '옥성', '케미칼']):
                        m = re.search(r'M/s\.\s*([^-\n\r]+)', txt)
                        if m:
                            supp = m.group(1).strip()
                        else:
                            supp = txt.split('\n')[0].replace('RM SUPPLIER Name :', '').replace('RM Supplier -', '').strip()
                        break

            m_label = default_months[m_idx] if m_idx < len(default_months) else f"{m_idx+1}차"
            records.append({
                '법인': sub_default,
                'std_category': std_cat,
                'material_name': std_mat,
                '월': m_label,
                '단가': p_val,
                '통화': curr,
                '공급사': supp
            })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 5. App 실행 & 마스터 데이터 적재
# ---------------------------------------------------------
init_db()

df_db = load_data_from_db()
if df_db.empty:
    for default_name in ["[인도법인]법인별 원소재 가격 추이_260930_KR-EN R1.xlsx", "raw_material_sample.xlsx"]:
        if os.path.exists(default_name):
            init_records = parse_raw_material_excel(default_name)
            save_records_to_db(init_records)
            df_db = load_data_from_db()
            break

# ---------------------------------------------------------
# 6. 사이드바: 법인 / 원소재 재질 드롭다운 / 기간 필터
# ---------------------------------------------------------
st.sidebar.markdown("## ⚙️ 시세 조회 설정")

subs_options = ["전체 법인 (비교)", "한국", "인도", "중국"]
selected_sub_filter = st.sidebar.selectbox("🏢 법인 조회 범위", subs_options, index=0)

available_mats = [m for m in STANDARDIZED_MATERIALS if m in df_db['material_name'].unique()]
mat_options = ["전체 원소재 현황"] + (available_mats if available_mats else STANDARDIZED_MATERIALS)

selected_material = st.sidebar.selectbox(
    "📌 원소재 재질 선택",
    options=mat_options,
    index=0
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 기간 설정")
month_order = {'1월': 1, '2월': 2, '3월': 3, '4월': 4, '5월': 5, '6월': 6,
               '7월': 7, '8월': 8, '9월': 9, '10월': 10, '11월': 11, '12월': 12}

all_months = sorted(df_db['month'].unique(), key=lambda x: month_order.get(x, 99)) if not df_db.empty else ['7월', '8월', '9월']
selected_months = st.sidebar.multiselect("조회 대상 월 선택", options=all_months, default=all_months)

working_df = df_db.copy()
if selected_sub_filter != "전체 법인 (비교)":
    working_df = working_df[working_df['subsidiary'] == selected_sub_filter]

if selected_months:
    working_df = working_df[working_df['month'].isin(selected_months)]

# ---------------------------------------------------------
# 7. 메인 화면: 엑셀 복사/붙여넣기 창
# ---------------------------------------------------------
st.title("📊 글로벌 원소재 단가 누적 Trend 대시보드")

with st.expander("📋 사내 엑셀 단가표 복사·붙여넣기 (클릭하여 열기)", expanded=False):
    st.info("💡 **사용법**: 엑셀 표를 복사(Ctrl+C) 후 아래에 붙여넣기(Ctrl+V)하세요. 루피(₹)나 원화(₩) 기호가 섞여 있어도 자동으로 정확한 단가만 인식됩니다.")
    p_sub = st.selectbox("붙여넣을 대상 법인", ["인도", "한국", "중국"], index=0)
    pasted_text = st.text_area("엑셀 붙여넣기 창", height=100, placeholder="엑셀 복사 데이터를 여기에 붙여넣으세요...")
    if st.button("🚀 데이터 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_df = parse_smart_excel_clipboard(pasted_text, sub_default=p_sub)
            if not parsed_df.empty:
                cnt = save_records_to_db(parsed_df)
                st.success(f"총 {cnt}건의 원소재 시세 데이터가 정제되어 DB에 정상 누적되었습니다.")
                st.rerun()
            else:
                st.error("지정된 5대 표준 원소재(Cu, ALDC12, 사출수지 등) 단가를 찾을 수 없습니다.")
        else:
            st.warning("붙여넣은 내용이 없습니다.")

if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 안내] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 데이터 접수 시 복사·붙여넣기하면 자동 누적됩니다.")

# ---------------------------------------------------------
# 8. 재질별 개별 그래프 & 법인별 비교 렌더링
# ---------------------------------------------------------
st.subheader(f"📈 {selected_material} 단가 Trend (재질별 개별 그래프 & 법인별 비교)")

if selected_material == "전체 원소재 현황":
    materials_to_show = [m for m in STANDARDIZED_MATERIALS if m in working_df['material_name'].unique()]
else:
    materials_to_show = [selected_material]

sub_color_map = {
    '한국': '#1f77b4',  # Blue
    '인도': '#2ca02c',  # Green
    '중국': '#ff7f0e'   # Orange
}

if materials_to_show and not working_df.empty:
    for i in range(0, len(materials_to_show), 2):
        row_mats = materials_to_show[i:i+2]
        cols = st.columns(len(row_mats))
        
        for idx, mat_name in enumerate(row_mats):
            mat_data = working_df[working_df['material_name'] == mat_name].dropna(subset=['price']).copy()
            
            with cols[idx]:
                if not mat_data.empty:
                    mat_data['month_rank'] = mat_data['month'].map(lambda x: month_order.get(x, 99))
                    mat_data = mat_data.sort_values(by=['month_rank', 'subsidiary'])
                    
                    mat_data['fmt_price'] = mat_data.apply(lambda r: format_price_label(r['price'], r['currency']), axis=1)
                    mat_data['disp_supplier'] = mat_data['supplier'].apply(lambda s: s if s else "-")
                    
                    fig = px.line(
                        mat_data,
                        x='month',
                        y='price',
                        color='subsidiary',
                        color_discrete_map=sub_color_map,
                        markers=True,
                        title=f"📌 {mat_name}",
                        custom_data=['subsidiary', 'fmt_price', 'disp_supplier']
                    )
                    
                    fig.update_traces(
                        line=dict(width=2.5),
                        marker=dict(size=8),
                        hovertemplate="<b>%{x}</b><br>%{customdata[0]}: <b>%{customdata[1]}</b><br>공급처: %{customdata[2]}<extra></extra>"
                    )
                    
                    fig.update_layout(
                        xaxis_title="",
                        yaxis_title="단가",
                        hovermode="x unified",
                        xaxis=dict(type='category', categoryorder='array', categoryarray=all_months),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=-0.3,
                            xanchor="center",
                            x=0.5,
                            title=""
                        ),
                        margin=dict(l=30, r=30, t=50, b=50),
                        height=350
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.markdown(f"**📌 {mat_name}**")
                    st.info("해당 기간에 등록된 유효 단가 데이터가 없습니다.")
else:
    st.info("조회할 유효 원소재 데이터가 없습니다.")

# ---------------------------------------------------------
# 9. 세부 단가 피벗 테이블 및 MoM 증감율
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📋 세부 단가 및 MoM 변동 분석표")

display_table_df = working_df[working_df['material_name'].isin(materials_to_show)] if materials_to_show else pd.DataFrame()

if not display_table_df.empty:
    pivot_df = display_table_df.pivot_table(
        index=['std_category', 'material_name', 'subsidiary', 'currency', 'supplier'],
        columns='month',
        values='price',
        aggfunc='first'
    ).reset_index()

    active_m_cols = [m for m in all_months if m in pivot_df.columns and m in selected_months]
    format_dict = {m: '{:,.2f}' for m in active_m_cols}
    
    if len(active_m_cols) >= 2:
        prev_m, curr_m = active_m_cols[-2], active_m_cols[-1]
        diff_col = f'MoM 증감 ({curr_m} vs {prev_m})'
        rate_col = 'MoM 증감률(%)'
        pivot_df[diff_col] = pivot_df[curr_m] - pivot_df[prev_m]
        pivot_df[rate_col] = ((pivot_df[curr_m] - pivot_df[prev_m]) / pivot_df[prev_m]) * 100
        format_dict[diff_col] = '{:+,.2f}'
        format_dict[rate_col] = '{:+.2f}%'

    st.dataframe(pivot_df.style.format(format_dict, na_rep="-"), use_container_width=True)

    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
        df_db.to_excel(writer, sheet_name='DB_전체누적데이터', index=False)
        pivot_df.to_excel(writer, sheet_name='피벗요약', index=False)

    st.download_button(
        label="📥 누적 원소재 마스터 엑셀 다운로드 (백업)",
        data=out_buf.getvalue(),
        file_name=f"원소재_시세_누적마스터_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/octet-stream"
    )
