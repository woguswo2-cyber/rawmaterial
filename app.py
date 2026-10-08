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
    page_title="글로벌 원소재 시세 Trend 대시보드",
    page_icon="📈",
    layout="wide"
)

DB_PATH = "raw_materials_trend.db"

# ---------------------------------------------------------
# 1. SQLite Database Storage Engine (누적 관리)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS material_prices (
            subsidiary TEXT,
            category TEXT,
            material TEXT,
            grade TEXT,
            item_name TEXT,
            month TEXT,
            price REAL,
            currency TEXT,
            supplier TEXT,
            updated_at TEXT,
            PRIMARY KEY (subsidiary, item_name, month)
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
                subsidiary, category, material, grade, item_name, month, price, currency, supplier, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(subsidiary, item_name, month) DO UPDATE SET
                category=excluded.category,
                material=excluded.material,
                grade=excluded.grade,
                price=excluded.price,
                currency=excluded.currency,
                supplier=excluded.supplier,
                updated_at=excluded.updated_at
        ''', (
            str(row['법인']), str(row['구분']), str(row['재질']), str(row['GRADE']),
            str(row['원소재품목']), str(row['월']), p_val, str(row['통화']),
            str(row.get('공급사', '')), now_str
        ))
        count += 1
    conn.commit()
    conn.close()
    return count

def load_data_from_db():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query("SELECT * FROM material_prices ORDER BY category, material, item_name, month", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    return df

# ---------------------------------------------------------
# 2. 사내 엑셀 복사/붙여넣기 전용 파서 (인도/한국 엑셀 완벽 대응)
# ---------------------------------------------------------
def parse_excel_clipboard(pasted_text, sub_default="한국"):
    lines = [l for l in pasted_text.strip().splitlines() if l.strip()]
    if not lines:
        return pd.DataFrame()

    sep = '\t' if '\t' in pasted_text else ','
    rows = [[c.strip() for c in l.split(sep)] for l in lines]

    # 1) '7월', '8월', '9월' 등 월 헤더가 있는 행 식별
    month_row_idx = -1
    month_indices = []

    for r_idx, row in enumerate(rows[:6]):
        m_matches = []
        for c_idx, cell in enumerate(row):
            m = re.search(r'(\d{1,2}월|\d{1,2}Q|\d{4}-\d{2})', cell)
            if m:
                m_matches.append((c_idx, m.group(1)))
        if len(m_matches) >= 2:
            month_row_idx = r_idx
            month_indices = m_matches
            break

    if month_row_idx == -1:
        return pd.DataFrame()

    # 데이터 시작 행 (단가, Unit Price 등이 있는 행 스킵)
    data_start_idx = month_row_idx + 1
    if data_start_idx < len(rows) and any('단가' in c or 'Price' in c for c in rows[data_start_idx]):
        data_start_idx += 1

    records = []
    curr_cat = ""
    curr_mat = ""

    for r in range(data_start_idx, len(rows)):
        row = rows[r]
        if not any(row):
            continue

        # 구분
        c0 = row[0] if len(row) > 0 else ""
        if c0:
            clean_c0 = c0.split('/')[0].split('\n')[0].strip()
            if clean_c0:
                curr_cat = clean_c0

        # 재질
        c1 = row[1] if len(row) > 1 else ""
        if c1:
            clean_c1 = c1.split('/')[0].split('\n')[0].strip()
            if clean_c1:
                curr_mat = clean_c1
        elif not c1 and curr_cat in ['Cu(Kg)', 'LME', 'TTS', 'Sn', '황동원소재', 'Cu', '황동']:
            curr_mat = curr_cat

        # GRADE
        c2 = row[2] if len(row) > 2 else ""
        grade = c2 if c2 and c2 not in ['nan', '-', ''] else 'STD'

        # 원소재 품목 표시명 (예: Cu(Kg), 알루미늄 ALDC12종, 사출원재료 PP-TD20 (JF1512))
        if curr_mat == curr_cat:
            item_display = curr_mat
        else:
            item_display = f"{curr_cat} {curr_mat}".strip()
            
        if grade != 'STD':
            item_display += f" ({grade})"

        for c_idx, m_name in month_indices:
            price_str = row[c_idx] if c_idx < len(row) else ""
            clean_p = price_str.replace(',', '').replace(' ', '')
            try:
                p_val = float(clean_p)
            except Exception:
                p_val = None

            curr_str = row[c_idx+1] if c_idx+1 < len(row) else "KRW"
            if not curr_str or any(char.isdigit() for char in curr_str):
                curr_str = "KRW"

            supp_str = row[c_idx+2] if c_idx+2 < len(row) else ""

            records.append({
                '법인': sub_default,
                '구분': curr_cat if curr_cat else '원소재',
                '재질': curr_mat if curr_mat else '주요소재',
                'GRADE': grade,
                '원소재품목': item_display,
                '월': m_name,
                '단가': p_val,
                '통화': curr_str,
                '공급사': supp_str
            })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 3. 엑셀 파일(.xlsx) 전체 시트 자동 파싱 엔진
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
        curr_cat = ""
        curr_mat = ""
        for r in range(5, len(df_summary)):
            c = df_summary.iloc[r, 1]
            if pd.notna(c) and str(c).strip() != "":
                curr_cat = str(c).split('\n')[0].split('/')[0].strip()
            m = df_summary.iloc[r, 2]
            if pd.notna(m) and str(m).strip() != "":
                curr_mat = str(m).split('\n')[0].split('/')[0].strip()
            elif pd.isna(m):
                curr_mat = curr_cat

            grade = str(df_summary.iloc[r, 3]).strip() if pd.notna(df_summary.iloc[r, 3]) else 'STD'
            if grade in ['nan', '-', '']:
                grade = 'STD'

            if curr_mat == curr_cat:
                item_display = curr_mat
            else:
                item_display = f"{curr_cat} {curr_mat}".strip()
            if grade != 'STD':
                item_display += f" ({grade})"

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
                        '구분': curr_cat,
                        '재질': curr_mat,
                        'GRADE': grade,
                        '원소재품목': item_display,
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
            grade = str(df_in.iloc[r, 3]).strip() if pd.notna(df_in.iloc[r, 3]) else 'STD'
            if grade in ['nan', '-', '']:
                continue

            item_display = f"{curr_cat} {curr_mat} ({grade})"

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
                    '구분': curr_cat,
                    '재질': curr_mat,
                    'GRADE': grade,
                    '원소재품목': item_display,
                    '월': month_name,
                    '단가': p_val,
                    '통화': str(curr).strip() if pd.notna(curr) else 'INR',
                    '공급사': str(supp).strip() if pd.notna(supp) else ''
                })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 4. App Execution & DB Initialization
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
# 5. 사이드바: 법인 선택 & 전체 현황 드롭다운 & 기간 설정
# ---------------------------------------------------------
st.sidebar.markdown("## ⚙️ 시세 조회 설정")

# 1) 법인 선택 (전체 법인 / 한국 / 인도 / 중국)
subs_options = ["전체 법인", "한국", "인도", "중국"]
selected_sub = st.sidebar.selectbox("🏢 법인 선택", subs_options, index=0)

# 법인 필터링
if not df_db.empty:
    if selected_sub != "전체 법인":
        sub_df = df_db[df_db['subsidiary'] == selected_sub]
    else:
        sub_df = df_db.copy()
else:
    sub_df = pd.DataFrame()

# 2) 원소재 현황 드롭다운 (1순위: 전체 원소재 현황, 2순위: 개별 품목)
item_options = ["전체 원소재 현황"]
if not sub_df.empty:
    available_items = sorted([str(i) for i in sub_df['item_name'].dropna().unique() if str(i).strip() != ''])
    item_options.extend(available_items)

selected_item = st.sidebar.selectbox(
    "📌 원소재 선택 (드롭다운)",
    options=item_options,
    index=0,
    help="'전체 원소재 현황'을 선택하면 Cu, AL, 사출원재료 등 모든 품목의 누적 트렌드를 한눈에 조회합니다."
)

# 3) 기간(월) 설정 필터
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

# 데이터 필터링 적용
if not sub_df.empty and selected_months:
    final_df = sub_df[sub_df['month'].isin(selected_months)].copy()
    if selected_item != "전체 원소재 현황":
        final_df = final_df[final_df['item_name'] == selected_item]
else:
    final_df = pd.DataFrame()

# ---------------------------------------------------------
# 6. 메인 화면: 데이터 입력 (Ctrl+C / Ctrl+V)
# ---------------------------------------------------------
st.title("📊 글로벌 원소재 단가 누적 Trend 분석 대시보드")

with st.expander("📋 사내 엑셀 단가표 복사·붙여넣기 (클릭하여 열기)", expanded=df_db.empty):
    st.info("💡 **사용법**: 사내 엑셀에서 단가 표 영역(7월/8월/9월 헤더 포함)을 드래그하여 **Ctrl+C**한 뒤 아래에 **Ctrl+V**하세요.")
    p_sub = st.selectbox("붙여넣을 대상 법인", ["한국", "인도", "중국"], index=0)
    pasted_text = st.text_area("엑셀 표 붙여넣기 창", height=130, placeholder="엑셀에서 복사한 단가 표를 여기에 붙여넣기하세요...")
    
    if st.button("🚀 데이터 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_df = parse_excel_clipboard(pasted_text, sub_default=p_sub)
            if not parsed_df.empty:
                cnt = save_records_to_db(parsed_df)
                st.success(f"성공! 총 {cnt}건의 데이터가 정제되어 DB에 누적 저장되었습니다.")
                st.rerun()
            else:
                st.error("데이터 파싱 실패: 복사한 영역에 '7월', '8월' 등의 월 헤더가 포함되었는지 확인해주세요.")
        else:
            st.warning("붙여넣은 내용이 없습니다.")

# 중국 법인 미접수 안내 배너
if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 안내] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 데이터 접수 시 복사·붙여넣기로 추가하면 통합 반영됩니다.")

# ---------------------------------------------------------
# 7. Trend 꺾은선 그래프 (순수 누적 트렌드 전용)
# ---------------------------------------------------------
st.subheader(f"📈 [{selected_sub}] {selected_item} 단가 누적 Trend")

valid_plot_df = final_df.dropna(subset=['price']).copy() if not final_df.empty else pd.DataFrame()

if not valid_plot_df.empty:
    valid_plot_df['month_rank'] = valid_plot_df['month'].map(lambda x: month_order.get(x, 99))
    valid_plot_df = valid_plot_df.sort_values(by=['month_rank', 'subsidiary', 'item_name'])

    if selected_sub == "전체 법인":
        valid_plot_df['line_label'] = valid_plot_df.apply(
            lambda r: f"[{r['subsidiary']}] {r['item_name']} ({r['currency']})", axis=1
        )
    else:
        valid_plot_df['line_label'] = valid_plot_df.apply(
            lambda r: f"{r['item_name']} ({r['currency']})", axis=1
        )

    fig = px.line(
        valid_plot_df,
        x='month',
        y='price',
        color='line_label',
        markers=True,
        hover_data={
            'subsidiary': True,
            'item_name': True,
            'currency': True,
            'supplier': True,
            'price': ':.2f',
            'month': False,
            'line_label': False
        }
    )

    fig.update_layout(
        xaxis_title="",  # 불필요한 라벨 제거
        yaxis_title="단가",
        hovermode="x unified",
        xaxis=dict(type='category', categoryorder='array', categoryarray=all_months),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.5,
            xanchor="center",
            x=0.5
        ),
        margin=dict(l=40, r=40, t=30, b=120)
    )
    fig.update_traces(line=dict(width=2.5), marker=dict(size=7))
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("조회할 단가 데이터가 없습니다. 상단에서 엑셀 표를 복사·붙여넣기하거나 사이드바의 필터를 조정해주세요.")

# ---------------------------------------------------------
# 8. 세부 단가 피벗 테이블 및 MoM 증감율
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📋 원소재별 세부 단가 및 MoM 변동 분석")

if not final_df.empty:
    pivot_df = final_df.pivot_table(
        index=['subsidiary', 'category', 'item_name', 'currency', 'supplier'],
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
        pivot_df.to_excel(writer, sheet_name='조회_피벗요약', index=False)

    st.download_button(
        label="📥 누적 원소재 마스터 엑셀 다운로드 (백업)",
        data=out_buf.getvalue(),
        file_name=f"원소재_시세_누적마스터_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/octet-stream"
    )
