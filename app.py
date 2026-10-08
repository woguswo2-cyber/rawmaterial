import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import re
import io
import os
from datetime import datetime

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="원소재 시세 Trend 모니터링 대시보드",
    page_icon="📊",
    layout="wide"
)

DB_PATH = "raw_materials_trend.db"

# ---------------------------------------------------------
# 원소재 표준 마스터 정의 (지정 분류 체계)
# ---------------------------------------------------------
def match_master_material(cat, mat, grade):
    text = f"{cat} {mat} {grade}".upper()
    
    # 3. 사출 원재료 세부 그레이드
    if "JF1512" in text: return "3. 사출 원재료", "PP-TD20 (JF1512)"
    if "JF1514" in text: return "3. 사출 원재료", "PP-TD+GF (JF1514G10)"
    if "JF5513" in text: return "3. 사출 원재료", "PP-GF30% (JF5513)"
    if "G-152" in text or "G152" in text: return "3. 사출 원재료", "PP-GF20 (G-152)"
    if "G-153B" in text or "G153B" in text: return "3. 사출 원재료", "PP-GF30 (G-153B)"
    if "JI350N" in text: return "3. 사출 원재료", "PP (JI350N)"
    if "JI350G" in text: return "3. 사출 원재료", "PP (JI350G)"
    if "KN333G30BL" in text or "KN333" in text: return "3. 사출 원재료", "PA66-GF30 (KN333G30BL)"
    if "KP213G30BL" in text or "KP213" in text: return "3. 사출 원재료", "PBT-GF30 (KP213G30BL)"
    if "LGP2200-BK" in text: return "3. 사출 원재료", "PP-GF20 (LGP2200-BK)"
    if "LGP2200-NP" in text: return "3. 사출 원재료", "PP-GF20 (LGP2200-NP)"
    if "LGP2100-NP" in text or "LGP2100" in text: return "3. 사출 원재료", "PP-GF10 (LGP2100-NP)"
    
    # 인도 로컬 사출 그레이드
    if "HT840IN" in text: return "3. 사출 원재료", "PP-TD20 (HT840IN BK)"
    if "HG940IN" in text: return "3. 사출 원재료", "PP-GF20 (HG940IN)"
    if "FZ1140D5" in text: return "3. 사출 원재료", "PPS (FZ1140D5 NAT)"
    if "ULTRADUR" in text: return "3. 사출 원재료", "PBT-GF30 (ULTRADUR B4300)"
    if "RYNITE" in text: return "3. 사출 원재료", "PET+GF35 (RYNITE 935)"
    if "ZYTEL" in text: return "3. 사출 원재료", "PA66+GF30 (ZYTEL)"
    
    # 4. 황동 원소재
    if "C1100" in text or "황동" in text or "BRASS" in text:
        return "4. 황동 원소재", "황동 C1100"
        
    # 2. ALDC 12
    if "ALDC" in text or "알루미늄" in text or "ALUMINUM" in text:
        return "2. ALDC 12", "ALDC 12"
        
    # 5. Sn
    if "SN" in text or "주석" in text:
        return "5. Sn", "Sn"
        
    # 1. Cu & 참고 지표
    if "LME" in text:
        return "6. 참고지표 (LME / TTS)", "LME 전기동"
    if "TTS" in text:
        return "6. 참고지표 (LME / TTS)", "환율 TTS"
    if "CU" in text or "구리" in text or "동" in text:
        return "1. Cu", "Cu"
        
    fallback_name = f"{mat} ({grade})" if grade and grade != 'STD' else (mat if mat else cat)
    return "기타 원소재", fallback_name

# ---------------------------------------------------------
# 1. SQLite Database Storage Engine (누적 관리)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='material_prices'")
    if cur.fetchone():
        cur.execute("PRAGMA table_info(material_prices)")
        cols = [c[1] for c in cur.fetchall()]
        if 'std_category' not in cols:
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
# 2. 스마트 엑셀 복사/붙여넣기 파서
# ---------------------------------------------------------
def parse_smart_excel_clipboard(pasted_text, sub_default="한국", default_months=None):
    if default_months is None:
        default_months = ["7월", "8월", "9월", "10월", "11월", "12월"]
        
    lines = [l for l in pasted_text.strip().splitlines() if l.strip()]
    if not lines:
        return pd.DataFrame()

    sep = '\t' if '\t' in pasted_text else ','
    rows = [[c.strip() for c in l.split(sep)] for l in lines]

    month_row_idx = -1
    month_indices = []
    for r_idx, row in enumerate(rows[:5]):
        m_matches = []
        for c_idx, cell in enumerate(row):
            m = re.search(r'(\d{1,2}월|\d{1,2}Q|\d{4}-\d{2})', cell)
            if m:
                m_matches.append((c_idx, m.group(1)))
        if len(m_matches) >= 2:
            month_row_idx = r_idx
            month_indices = m_matches
            break

    records = []

    # Case A: 월 헤더가 포함된 복사
    if month_row_idx != -1:
        data_start_idx = month_row_idx + 1
        if data_start_idx < len(rows) and any('단가' in c or 'Price' in c for c in rows[data_start_idx]):
            data_start_idx += 1

        curr_cat, curr_mat = "", ""
        for r in range(data_start_idx, len(rows)):
            row = rows[r]
            if not any(row):
                continue
            if row[0]:
                curr_cat = row[0].split('/')[0].split('\n')[0].strip()
            if len(row) > 1 and row[1]:
                curr_mat = row[1].split('/')[0].split('\n')[0].strip()
            elif not curr_mat and curr_cat:
                curr_mat = curr_cat

            grade = row[2] if len(row) > 2 and row[2] not in ['nan', '-', ''] else ''
            std_cat, std_mat = match_master_material(curr_cat, curr_mat, grade)

            for c_idx, m_name in month_indices:
                p_val = None
                curr_str = "KRW"
                supp_str = ""
                if c_idx < len(row):
                    clean_p = row[c_idx].replace(',', '').replace(' ', '')
                    try:
                        p_val = float(clean_p)
                    except Exception:
                        p_val = None
                if c_idx + 1 < len(row):
                    cand_c = row[c_idx + 1]
                    if cand_c and not any(char.isdigit() for char in cand_c):
                        curr_str = cand_c
                if c_idx + 2 < len(row):
                    supp_str = row[c_idx + 2]

                records.append({
                    '법인': sub_default,
                    'std_category': std_cat,
                    'material_name': std_mat,
                    '월': m_name,
                    '단가': p_val,
                    '통화': curr_str,
                    '공급사': supp_str
                })

    # Case B: 순수 데이터 블록만 복사한 경우
    else:
        for row in rows:
            if not any(row):
                continue
            first_num_idx = -1
            for i, cell in enumerate(row):
                clean = cell.replace(',', '').replace(' ', '')
                try:
                    val = float(clean)
                    if val > 0:
                        first_num_idx = i
                        break
                except Exception:
                    pass

            if first_num_idx == -1:
                continue

            name_tokens = [c.split('/')[0].strip() for c in row[:first_num_idx] if c.strip()]
            c_text = " ".join(name_tokens)
            std_cat, std_mat = match_master_material(c_text, "", "")

            val_part = row[first_num_idx:]
            idx = 0
            m_count = 0
            while idx < len(val_part):
                cell = val_part[idx]
                clean = cell.replace(',', '').replace(' ', '')
                try:
                    p_val = float(clean)
                    curr = "KRW"
                    supp = ""
                    if idx + 1 < len(val_part) and val_part[idx+1] in ['KRW', 'USD', 'INR', 'CNY']:
                        curr = val_part[idx+1]
                        if idx + 2 < len(val_part) and not any(c.isdigit() for c in val_part[idx+2]):
                            supp = val_part[idx+2]
                            idx += 3
                        else:
                            idx += 2
                    else:
                        idx += 1

                    m_label = default_months[m_count] if m_count < len(default_months) else f"{m_count+1}차"
                    records.append({
                        '법인': sub_default,
                        'std_category': std_cat,
                        'material_name': std_mat,
                        '월': m_label,
                        '단가': p_val,
                        '통화': curr,
                        '공급사': supp
                    })
                    m_count += 1
                except Exception:
                    idx += 1

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 3. 엑셀 파일(.xlsx) 자동 시드 파서
# ---------------------------------------------------------
def parse_raw_material_excel(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    records = []

    if '요약_Summary' in excel.sheet_names:
        df_summary = pd.read_excel(excel, sheet_name='요약_Summary', header=None)
        subs_map = {
            '한국': {'7월': 4, '8월': 8, '9월': 12, 'default_curr': 'KRW'},
            '중국': {'7월': 16, '8월': 20, '9월': 24, 'default_curr': 'CNY'},
            '인도': {'7월': 28, '8월': 32, '9월': 36, 'default_curr': 'INR'}
        }
        curr_cat, curr_mat = "", ""
        for r in range(5, len(df_summary)):
            c = df_summary.iloc[r, 1]
            if pd.notna(c) and str(c).strip() != "":
                curr_cat = str(c).split('\n')[0].split('/')[0].strip()
            m = df_summary.iloc[r, 2]
            if pd.notna(m) and str(m).strip() != "":
                curr_mat = str(m).split('\n')[0].split('/')[0].strip()
            elif pd.isna(m):
                curr_mat = curr_cat

            grade = str(df_summary.iloc[r, 3]).strip() if pd.notna(df_summary.iloc[r, 3]) else ''
            if grade in ['nan', '-']:
                grade = ''

            std_cat, std_mat = match_master_material(curr_cat, curr_mat, grade)

            for sub, m_info in subs_map.items():
                for month_name in ['7월', '8월', '9월']:
                    col_idx = m_info[month_name]
                    price = df_summary.iloc[r, col_idx]
                    curr = df_summary.iloc[r, col_idx + 1]
                    supp = df_summary.iloc[r, col_idx + 2]
                    try:
                        p_val = float(price) if pd.notna(price) else None
                    except Exception:
                        p_val = None

                    records.append({
                        '법인': sub,
                        'std_category': std_cat,
                        'material_name': std_mat,
                        '월': month_name,
                        '단가': p_val,
                        '통화': str(curr).strip() if pd.notna(curr) else m_info['default_curr'],
                        '공급사': str(supp).strip() if pd.notna(supp) else ''
                    })

    if 'HS India  Grades' in excel.sheet_names:
        df_in = pd.read_excel(excel, sheet_name='HS India  Grades', header=None)
        curr_cat = "사출원재료"
        curr_mat = ""
        for r in range(5, len(df_in)):
            m = df_in.iloc[r, 2]
            if pd.notna(m) and str(m).strip() != "":
                curr_mat = str(m).split('\n')[0].strip()
            grade = str(df_in.iloc[r, 3]).strip() if pd.notna(df_in.iloc[r, 3]) else ''
            if grade in ['nan', '-', '']:
                continue

            std_cat, std_mat = match_master_material(curr_cat, curr_mat, grade)

            for month_name, col_idx in [('7월', 28), ('8월', 32), ('9월', 36)]:
                price = df_in.iloc[r, col_idx]
                curr = df_in.iloc[r, col_idx + 1]
                supp = df_in.iloc[r, col_idx + 2]
                try:
                    p_val = float(price) if pd.notna(price) else None
                except Exception:
                    p_val = None

                records.append({
                    '법인': '인도',
                    'std_category': std_cat,
                    'material_name': std_mat,
                    '월': month_name,
                    '단가': p_val,
                    '통화': str(curr).strip() if pd.notna(curr) else 'INR',
                    '공급사': str(supp).strip() if pd.notna(supp) else ''
                })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 4. App Execution & DB Initial Load
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
# 5. 사이드바: 법인 / 원소재 재질 드롭다운 / 기간 필터
# ---------------------------------------------------------
st.sidebar.markdown("## ⚙️ 시세 조회 설정")

subs_options = ["전체 법인 (비교)", "한국", "인도", "중국"]
selected_sub_filter = st.sidebar.selectbox("🏢 법인 조회 범위", subs_options, index=0)

all_db_materials = sorted(df_db['material_name'].dropna().unique().tolist()) if not df_db.empty else []
mat_options = ["전체 원소재 현황"] + all_db_materials

selected_material = st.sidebar.selectbox(
    "📌 원소재 재질 선택",
    options=mat_options,
    index=0,
    help="'전체 원소재 현황'을 선택하면 지정된 재질별로 개별 그래프가 나열되며, 각 그래프 내에 법인별 단가 추이가 표기됩니다."
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📅 조회 기간 설정")

month_order = {'1월': 1, '2월': 2, '3월': 3, '4월': 4, '5월': 5, '6월': 6,
               '7월': 7, '8월': 8, '9월': 9, '10월': 10, '11월': 11, '12월': 12}

all_months = []
if not df_db.empty:
    raw_months = [m for m in df_db['month'].unique() if pd.notna(m)]
    all_months = sorted(raw_months, key=lambda x: month_order.get(x, 99))

if all_months:
    selected_months = st.sidebar.multiselect(
        "조회 대상 기간 선택",
        options=all_months,
        default=all_months
    )
else:
    selected_months = []

working_df = df_db.copy()
if selected_sub_filter != "전체 법인 (비교)":
    working_df = working_df[working_df['subsidiary'] == selected_sub_filter]

if selected_months:
    working_df = working_df[working_df['month'].isin(selected_months)]

# ---------------------------------------------------------
# 6. 메인 화면: 엑셀 복사/붙여넣기 창
# ---------------------------------------------------------
st.title("📊 글로벌 원소재 단가 누적 Trend 모니터링")

with st.expander("📋 사내 엑셀 단가표 복사·붙여넣기 (클릭하여 열기)", expanded=df_db.empty):
    st.info("💡 **사용법**: 사내 엑셀에서 단가 표 영역을 복사(Ctrl+C)한 후 아래에 붙여넣기(Ctrl+V)하세요.")
    p_sub = st.selectbox("붙여넣을 대상 법인", ["한국", "인도", "중국"], index=0)
    pasted_text = st.text_area("엑셀 표 붙여넣기 창", height=120, placeholder="엑셀 복사 데이터를 여기에 붙여넣으세요...")
    
    if st.button("🚀 데이터 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_df = parse_smart_excel_clipboard(pasted_text, sub_default=p_sub)
            if not parsed_df.empty:
                cnt = save_records_to_db(parsed_df)
                st.success(f"총 {cnt}건의 원소재 시세 데이터가 정제되어 DB에 누적 저장되었습니다.")
                st.rerun()
            else:
                st.error("데이터 파싱 실패: 복사한 데이터에 수치 단가가 포함되어 있는지 확인해주세요.")
        else:
            st.warning("붙여넣은 내용이 없습니다.")

if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 안내] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 데이터 접수 시 복사·붙여넣기하면 자동 누적됩니다.")

# ---------------------------------------------------------
# 7. 재질별 개별 그래프 렌더링 (각 그래프 안에 법인별 라인)
# ---------------------------------------------------------
st.subheader(f"📈 {selected_material} 단가 Trend (재질별 개별 그래프 & 법인별 비교)")

if selected_material == "전체 원소재 현황":
    materials_to_show = sorted(working_df['material_name'].dropna().unique().tolist())
else:
    materials_to_show = [selected_material]

sub_color_map = {
    '한국': '#1f77b4',  # 파랑
    '인도': '#2ca02c',  # 초록
    '중국': '#ff7f0e'   # 주황
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
                    
                    fig = px.line(
                        mat_data,
                        x='month',
                        y='price',
                        color='subsidiary',
                        color_discrete_map=sub_color_map,
                        markers=True,
                        title=f"📌 {mat_name}",
                        hover_data={
                            'subsidiary': True,
                            'currency': True,
                            'supplier': True,
                            'price': ':.2f',
                            'month': False
                        }
                    )
                    
                    fig.update_layout(
                        xaxis_title="",
                        yaxis_title="단가",
                        hovermode="x unified",
                        xaxis=dict(type='category', categoryorder='array', categoryarray=all_months),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=-0.35,
                            xanchor="center",
                            x=0.5,
                            title=""
                        ),
                        margin=dict(l=30, r=30, t=50, b=60),
                        height=360
                    )
                    fig.update_traces(line=dict(width=2.5), marker=dict(size=8))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.markdown(f"**📌 {mat_name}**")
                    st.info("해당 기간에 등록된 유효 단가 데이터가 없습니다.")
else:
    st.info("조회할 원소재 데이터가 없습니다. 상단에서 데이터를 붙여넣거나 필터 조건을 확인해주세요.")

# ---------------------------------------------------------
# 8. 세부 단가 피벗 테이블 및 MoM 증감율
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📋 원소재별 세부 단가 및 MoM 변동 분석")

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
