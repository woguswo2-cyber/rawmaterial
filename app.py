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
# 1. SQLite Database Storage Engine (자동 마이그레이션 포함)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    # 기존 구버전 스키마 체크 및 자동 갱신
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='material_prices'")
    if cur.fetchone():
        cur.execute("PRAGMA table_info(material_prices)")
        cols = [c[1] for c in cur.fetchall()]
        if 'item_name' not in cols:
            cur.execute("DROP TABLE material_prices")
            
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
        df = pd.read_sql_query("SELECT * FROM material_prices ORDER BY item_name, month", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    return df

# ---------------------------------------------------------
# 2. 초유연 엑셀 복사/붙여넣기 파서 (헤더 포함/미포함 모두 지원)
# ---------------------------------------------------------
def parse_smart_excel_clipboard(pasted_text, sub_default="한국", default_months=None):
    if default_months is None:
        default_months = ["7월", "8월", "9월", "10월", "11월", "12월"]
        
    lines = [l for l in pasted_text.strip().splitlines() if l.strip()]
    if not lines:
        return pd.DataFrame()

    sep = '\t' if '\t' in pasted_text else ','
    rows = [[c.strip() for c in l.split(sep)] for l in lines]

    # 1) '7월', '8월', '9월' 헤더가 포함되어 있는지 탐색
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

    # Case A: 헤더 행이 포함된 경우
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

            grade = row[2] if len(row) > 2 and row[2] not in ['nan', '-', ''] else 'STD'
            item_display = curr_mat if curr_mat == curr_cat else f"{curr_cat} {curr_mat}".strip()
            if grade != 'STD':
                item_display += f" ({grade})"

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
                    '구분': curr_cat if curr_cat else '원소재',
                    '재질': curr_mat if curr_mat else '원소재',
                    'GRADE': grade,
                    '원소재품목': item_display,
                    '월': m_name,
                    '단가': p_val,
                    '통화': curr_str,
                    '공급사': supp_str
                })

    # Case B: 헤더 없이 순수 데이터 행만 복사한 경우 (현재 첨부 이미지 상황 대응)
    else:
        for row in rows:
            if not any(row):
                continue
            # 첫 번째 수치 단가 위치 탐색
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
            if len(name_tokens) == 1:
                item_display = name_tokens[0]
            elif len(name_tokens) >= 2:
                item_display = f"{name_tokens[0]} ({name_tokens[1]})"
            else:
                item_display = "원소재"

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
                        '구분': '원소재',
                        '재질': item_display,
                        'GRADE': 'STD',
                        '원소재품목': item_display,
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

            grade = str(df_summary.iloc[r, 3]).strip() if pd.notna(df_summary.iloc[r, 3]) else 'STD'
            if grade in ['nan', '-', '']:
                grade = 'STD'

            item_display = curr_mat if curr_mat == curr_cat else f"{curr_cat} {curr_mat}".strip()
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
# 5. 사이드바: 법인 / 원소재(전체 및 개별) / 기간 설정
# ---------------------------------------------------------
st.sidebar.markdown("## ⚙️ 시세 조회 설정")

subs_options = ["전체 법인", "한국", "인도", "중국"]
selected_sub = st.sidebar.selectbox("🏢 법인 선택", subs_options, index=0)

if not df_db.empty:
    if selected_sub != "전체 법인":
        sub_df = df_db[df_db['subsidiary'] == selected_sub]
    else:
        sub_df = df_db.copy()
else:
    sub_df = pd.DataFrame()

# 1) 전체 원소재 현황 및 개별 품목 드롭다운
item_options = ["전체 원소재 현황"]
if not sub_df.empty:
    available_items = sorted([str(i) for i in sub_df['item_name'].dropna().unique() if str(i).strip() != ''])
    item_options.extend(available_items)

selected_item = st.sidebar.selectbox(
    "📌 원소재 선택 (드롭다운)",
    options=item_options,
    index=0,
    help="'전체 원소재 현황' 선택 시 Cu, AL, 사출수지 등 모든 소재의 누적 트렌드가 일괄 표시됩니다."
)

# 2) 기간(월) 설정
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

if not sub_df.empty and selected_months:
    final_df = sub_df[sub_df['month'].isin(selected_months)].copy()
    if selected_item != "전체 원소재 현황":
        final_df = final_df[final_df['item_name'] == selected_item]
else:
    final_df = pd.DataFrame()

# ---------------------------------------------------------
# 6. 메인 화면: 엑셀 복사/붙여넣기 창
# ---------------------------------------------------------
st.title("📊 글로벌 원소재 단가 누적 Trend 분석 대시보드")

with st.expander("📋 사내 엑셀 단가표 복사·붙여넣기 (클릭하여 열기)", expanded=df_db.empty):
    st.info("💡 **사용법**: 엑셀에서 표 영역을 복사(Ctrl+C)한 후 아래에 붙여넣기(Ctrl+V)하세요. 상단 헤더 포함 여부와 상관없이 자동 분석됩니다.")
    p_sub = st.selectbox("붙여넣을 대상 법인", ["한국", "인도", "중국"], index=0)
    pasted_text = st.text_area("엑셀 표 붙여넣기 창", height=130, placeholder="엑셀에서 복사한 영역을 여기에 붙여넣으세요...")
    
    if st.button("🚀 데이터 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_df = parse_smart_excel_clipboard(pasted_text, sub_default=p_sub)
            if not parsed_df.empty:
                cnt = save_records_to_db(parsed_df)
                st.success(f"성공! 총 {cnt}건의 데이터가 정제되어 DB에 누적 저장되었습니다.")
                st.rerun()
            else:
                st.error("데이터 파싱 실패: 수치 단가 데이터가 포함되어 있는지 확인해주세요.")
        else:
            st.warning("붙여넣은 내용이 없습니다.")

if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 안내] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 데이터 접수 시 복사·붙여넣기하면 자동 누적됩니다.")

# ---------------------------------------------------------
# 7. Trend 꺾은선 그래프 (월별 순수 누적 트렌드)
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
        xaxis_title="",
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
    st.info("조회할 단가 데이터가 없습니다. 상단에서 엑셀 표를 복사·붙여넣기하거나 사이드바 필터를 확인해주세요.")

# ---------------------------------------------------------
# 8. 세부 단가 피벗 테이블 및 MoM 증감율
# ---------------------------------------------------------
st.markdown("---")
st.subheader("📋 원소재별 세부 단가 및 MoM 변동 분석")

if not final_df.empty:
    pivot_df = final_df.pivot_table(
        index=['subsidiary', 'item_name', 'currency', 'supplier'],
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
