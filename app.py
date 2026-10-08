import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import csv
import io
import os
from datetime import datetime

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="글로벌 법인별 원소재 단가 Trend 분석기",
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
            month TEXT,
            price REAL,
            currency TEXT,
            supplier TEXT,
            memo TEXT,
            updated_at TEXT,
            PRIMARY KEY (subsidiary, category, material, grade, month)
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
            p_val = float(row['price']) if pd.notna(row['price']) else None
        except Exception:
            p_val = None
            
        cur.execute('''
            INSERT INTO material_prices (
                subsidiary, category, material, grade, month, price, currency, supplier, memo, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(subsidiary, category, material, grade, month) DO UPDATE SET
                price=excluded.price,
                currency=excluded.currency,
                supplier=excluded.supplier,
                memo=excluded.memo,
                updated_at=excluded.updated_at
        ''', (
            str(row['subsidiary']), str(row['category']), str(row['material']), str(row['grade']),
            str(row['month']), p_val, str(row['currency']), str(row.get('supplier', '')),
            str(row.get('memo', '')), now_str
        ))
        count += 1
    conn.commit()
    conn.close()
    return count

def load_data_from_db():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query("SELECT * FROM material_prices ORDER BY category, material, grade, month", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    return df

# ---------------------------------------------------------
# 2. 초강력 엑셀 복사/붙여넣기 파싱 엔진 (ParserError 원천 방지)
# ---------------------------------------------------------
def robust_parse_clipboard(pasted_text, sub_default="한국", curr_default="KRW"):
    if not pasted_text or not pasted_text.strip():
        return pd.DataFrame()

    lines = [line.strip() for line in pasted_text.strip().splitlines() if line.strip()]
    if not lines:
        return pd.DataFrame()

    # 탭 / 쉼표 구분자 자동 감지
    first_few = lines[:min(5, len(lines))]
    tab_count = sum(l.count('\t') for l in first_few)
    comma_count = sum(l.count(',') for l in first_few)
    sep = '\t' if tab_count >= comma_count else ','

    # 엑셀 셀 내 줄바꿈, 따옴표, 불규칙 열 길이 완벽 대응 파서
    try:
        reader = csv.reader(io.StringIO(pasted_text), delimiter=sep)
        raw_rows = [row for row in reader if any(cell.strip() for cell in row)]
    except Exception:
        raw_rows = [[c.strip() for c in l.split(sep)] for l in lines]

    if not raw_rows:
        return pd.DataFrame()

    # 제목 줄(헤더) 행 자동 탐색
    header_idx = 0
    for idx, row in enumerate(raw_rows[:6]):
        row_str = " ".join(row)
        if any(keyword in row_str for keyword in ['재질', 'GRADE', 'Grade', '월', '단가', 'Material', 'Price', '구분']):
            header_idx = idx
            break

    headers = [c.strip() for c in raw_rows[header_idx]]
    data_rows = raw_rows[header_idx + 1:]
    
    if not data_rows:
        return pd.DataFrame()

    max_len = max(len(headers), max(len(r) for r in data_rows))
    
    while len(headers) < max_len:
        headers.append(f"열_{len(headers)+1}")
    for i in range(len(headers)):
        if not headers[i]:
            headers[i] = f"열_{i+1}"

    padded_data = []
    for r in data_rows:
        r_padded = [str(cell).strip() for cell in r]
        while len(r_padded) < max_len:
            r_padded.append('')
        padded_data.append(r_padded[:max_len])

    df = pd.DataFrame(padded_data, columns=headers[:max_len])

    # 시세 월 컬럼 vs 부가정보 컬럼 자동 분별
    month_cols = []
    meta_cols = []
    for col in df.columns:
        col_clean = str(col).strip()
        if any(m in col_clean for m in ['월', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', '202']):
            month_cols.append(col)
        elif any(k in col_clean for k in ['구분', '재질', 'GRADE', 'Grade', '품목', 'Material', '공급사', '통화', '비고', 'Memo', '법인']):
            meta_cols.append(col)
        else:
            clean_s = df[col].astype(str).str.replace(',', '').str.replace(' ', '')
            num_ratio = pd.to_numeric(clean_s, errors='coerce').notna().mean()
            if num_ratio > 0.3:
                month_cols.append(col)
            else:
                meta_cols.append(col)

    if not month_cols:
        for col in reversed(df.columns):
            if col not in meta_cols:
                month_cols.append(col)
                break

    records = []
    curr_cat = ""
    curr_mat = ""

    for _, row in df.iterrows():
        # 구분
        raw_cat = ""
        for k in ['구분', 'Category']:
            if k in row and str(row[k]).strip() != "":
                raw_cat = str(row[k]).split('\n')[0].strip()
                break
        if raw_cat:
            curr_cat = raw_cat

        # 재질
        raw_mat = ""
        for k in ['재질', 'Material', '품목']:
            if k in row and str(row[k]).strip() != "":
                raw_mat = str(row[k]).split('\n')[0].strip()
                break
        if raw_mat:
            curr_mat = raw_mat
        elif not curr_mat and curr_cat:
            curr_mat = curr_cat

        # Grade
        grade = "STD"
        for k in ['GRADE', 'Grade', '그레이드']:
            if k in row and str(row[k]).strip() not in ['', 'nan', 'None']:
                grade = str(row[k]).strip()
                break

        # 법인
        sub = sub_default
        for k in ['법인', 'Subsidiary']:
            if k in row and str(row[k]).strip() != "":
                sub = str(row[k]).strip()
                break

        # 통화
        curr = curr_default
        for k in ['통화', '결제통화', 'Currency']:
            if k in row and str(row[k]).strip() != "":
                curr = str(row[k]).strip()
                break

        # 공급사
        supp = ""
        for k in ['공급사', '공급사명', 'Supplier']:
            if k in row and str(row[k]).strip() not in ['', 'nan']:
                supp = str(row[k]).strip()
                break

        # 비고
        memo = ""
        for k in ['비고', 'Memo']:
            if k in row and str(row[k]).strip() not in ['', 'nan']:
                memo = str(row[k]).strip()
                break

        for m in month_cols:
            val = row[m]
            try:
                p_val = float(str(val).replace(',', '').strip())
            except Exception:
                p_val = np.nan

            records.append({
                'subsidiary': sub,
                'category': curr_cat if curr_cat else '원소재',
                'material': curr_mat if curr_mat else '주요품목',
                'grade': grade,
                'month': str(m).strip(),
                'price': p_val,
                'currency': curr,
                'supplier': supp,
                'memo': memo
            })

    return pd.DataFrame(records)

# ---------------------------------------------------------
# 3. 엑셀 파일(.xlsx) 전체 시트 파싱 엔진
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
                curr_cat = str(c).split('\n')[0].replace(' / ', '/').strip()
            m = df_summary.iloc[r, 2]
            if pd.notna(m) and str(m).strip() != "":
                curr_mat = str(m).split('\n')[0].strip()
            elif pd.isna(m):
                curr_mat = curr_cat
                
            grade = str(df_summary.iloc[r, 3]).strip() if pd.notna(df_summary.iloc[r, 3]) else 'STD'
            
            for sub, m_info in subs_map.items():
                for month_name in ['7월', '8월', '9월']:
                    col_idx = m_info[month_name]
                    price = df_summary.iloc[r, col_idx]
                    curr = df_summary.iloc[r, col_idx + 1]
                    supp = df_summary.iloc[r, col_idx + 2]
                    memo = df_summary.iloc[r, col_idx + 3]
                    try:
                        p_val = float(price) if pd.notna(price) else None
                    except Exception:
                        p_val = None
                    records.append({
                        'subsidiary': sub,
                        'category': curr_cat,
                        'material': curr_mat,
                        'grade': grade,
                        'month': month_name,
                        'price': p_val,
                        'currency': str(curr).strip() if pd.notna(curr) else m_info['default_curr'],
                        'supplier': str(supp).strip() if pd.notna(supp) else '',
                        'memo': str(memo).strip() if pd.notna(memo) else ''
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
            if grade == 'nan' or not grade:
                continue
            for month_name, col_idx in [('7월', 28), ('8월', 32), ('9월', 36)]:
                price = df_in.iloc[r, col_idx]
                curr = df_in.iloc[r, col_idx + 1]
                supp = df_in.iloc[r, col_idx + 2]
                memo = df_in.iloc[r, col_idx + 3]
                try:
                    p_val = float(price) if pd.notna(price) else None
                except Exception:
                    p_val = None
                records.append({
                    'subsidiary': '인도',
                    'category': curr_cat,
                    'material': curr_mat,
                    'grade': grade,
                    'month': month_name,
                    'price': p_val,
                    'currency': str(curr).strip() if pd.notna(curr) else 'INR',
                    'supplier': str(supp).strip() if pd.notna(supp) else '',
                    'memo': str(memo).strip() if pd.notna(memo) else ''
                })
                
    return pd.DataFrame(records)

# ---------------------------------------------------------
# 4. App Execution & Initial Data Load
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
# 5. UI: 데이터 입력 (붙여넣기 vs 업로드)
# ---------------------------------------------------------
st.title("📊 글로벌 법인별 원소재 단가 Trend 분석기")
st.caption("사출수지, 비철금속(Cu, Al, 황동) 월별 시세 추이 모니터링 & DB 누적 관리")

input_tab1, input_tab2 = st.tabs(["📋 엑셀 데이터 직접 붙여넣기 (사내망/보안 환경용)", "📁 엑셀 파일 업로드 (모바일/기타)"])

with input_tab1:
    st.info("💡 **간편 사용법**: 사내 엑셀에서 단가 표 영역을 드래그하여 **복사(Ctrl+C)**한 후, 아래 입력창에 **붙여넣기(Ctrl+V)**하세요. 병합 셀이나 줄바꿈이 있어도 자동 정제됩니다.")
    p_col1, p_col2 = st.columns(2)
    with p_col1:
        paste_sub = st.selectbox("기본 법인 지정", ["한국", "인도", "중국"], index=0)
    with p_col2:
        paste_curr = st.selectbox("결제 통화 지정", ["KRW", "INR", "USD", "CNY"], index=0)
        
    pasted_text = st.text_area(
        "엑셀 표 붙여넣기 창",
        placeholder="엑셀에서 복사(Ctrl+C)한 표를 여기에 붙여넣기(Ctrl+V)하세요...",
        height=140
    )
    
    if st.button("🚀 붙여넣은 데이터로 즉시 차트 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_paste_df = robust_parse_clipboard(pasted_text, sub_default=paste_sub, curr_default=paste_curr)
            if not parsed_paste_df.empty:
                saved = save_records_to_db(parsed_paste_df)
                st.success(f"총 {saved}건의 시세 데이터가 정제되어 누적 DB에 정상 저장되었습니다.")
                df_db = load_data_from_db()
            else:
                st.error("데이터 파싱에 실패했습니다. 복사한 영역에 열 제목이나 수치가 포함되어 있는지 확인해주세요.")
        else:
            st.warning("붙여넣은 내용이 없습니다. 엑셀에서 복사 후 입력해주세요.")

with input_tab2:
    uploaded_file = st.file_uploader("최신 원소재 엑셀 파일 (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        parsed_file_df = parse_raw_material_excel(uploaded_file)
        saved = save_records_to_db(parsed_file_df)
        st.success(f"엑셀 파일에서 총 {saved}건 누적 저장 완료!")
        df_db = load_data_from_db()

# 중국 법인 미접수 안내 배너
if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 안내] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 데이터 접수 시 복사·붙여넣기로 추가하면 통합 반영됩니다.")

# ---------------------------------------------------------
# 6. 사이드바: 재질별 드롭다운 및 필터 설정
# ---------------------------------------------------------
st.sidebar.markdown("## 🔍 시세 조회 조건")

selected_material = "전체 재질"
selected_subs = []
selected_grades = []
final_df = pd.DataFrame()

if not df_db.empty:
    material_list = sorted([m for m in df_db['material'].dropna().unique() if str(m).strip() != ''])
    selected_material = st.sidebar.selectbox(
        "📌 재질(Material) 선택",
        options=["전체 재질"] + material_list,
        index=0
    )
    
    subs_list = list(df_db['subsidiary'].unique())
    selected_subs = st.sidebar.multiselect(
        "🏢 대상 법인 선택",
        options=subs_list,
        default=[s for s in ['한국', '인도'] if s in subs_list]
    )
    
    mask = df_db['subsidiary'].isin(selected_subs)
    if selected_material != "전체 재질":
        mask = mask & (df_db['material'] == selected_material)
    sub_filtered = df_db[mask]
    
    grade_list = sorted([g for g in sub_filtered['grade'].dropna().unique() if str(g).strip() != ''])
    selected_grades = st.sidebar.multiselect(
        "🏷️ GRADE 선택",
        options=grade_list,
        default=grade_list
    )
    
    final_df = sub_filtered[sub_filtered['grade'].isin(selected_grades)].copy()

st.markdown("---")

# ---------------------------------------------------------
# 7. Trend 꺾은선 그래프
# ---------------------------------------------------------
trend_title = f"[{selected_material}] 단가 변동 Trend (꺾은선 그래프)" if selected_material != "전체 재질" else "주요 원소재 단가 변동 Trend (꺾은선 그래프)"
st.subheader(f"📈 {trend_title}")

valid_plot_df = final_df.dropna(subset=['price']).copy() if not final_df.empty else pd.DataFrame()

if not valid_plot_df.empty:
    month_order = {'7월': 1, '8월': 2, '9월': 3, '10월': 4, '11월': 5, '12월': 6}
    valid_plot_df['month_rank'] = valid_plot_df['month'].map(lambda x: month_order.get(x, 99))
    valid_plot_df = valid_plot_df.sort_values(by=['month_rank', 'subsidiary', 'material', 'grade'])
    
    valid_plot_df['line_label'] = valid_plot_df.apply(
        lambda r: f"[{r['subsidiary']}] {r['material']} | {r['grade']} ({r['currency']})", axis=1
    )
    
    fig = px.line(
        valid_plot_df,
        x='month',
        y='price',
        color='line_label',
        markers=True,
        hover_data={
            'subsidiary': True,
            'material': True,
            'grade': True,
            'currency': True,
            'supplier': True,
            'memo': True,
            'price': ':.2f',
            'month': False,
            'line_label': False
        },
        title=f"원소재 월별 단가 변동 추이 ({selected_material})"
    )
    
    fig.update_layout(
        xaxis_title="기준월",
        yaxis_title="단가 (각 법인 통화 기준)",
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.5,
            xanchor="center",
            x=0.5
        ),
        margin=dict(l=40, r=40, t=50, b=100)
    )
    fig.update_traces(line=dict(width=2.5), marker=dict(size=8))
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("조회할 단가 데이터가 아직 없습니다. 상단에서 엑셀 표를 복사·붙여넣기하거나 파일을 업로드해주세요.")

# ---------------------------------------------------------
# 8. 세부 단가 피벗 테이블 및 MoM 증감율
# ---------------------------------------------------------
st.subheader("📋 세부 단가 피벗 테이블 및 MoM 분석")

if not final_df.empty:
    pivot_df = final_df.pivot_table(
        index=['subsidiary', 'category', 'material', 'grade', 'currency', 'supplier'],
        columns='month',
        values='price',
        aggfunc='first'
    ).reset_index()
    
    cols = [c for c in ['7월', '8월', '9월', '10월', '11월', '12월'] if c in pivot_df.columns]
    format_dict = {c: '{:,.2f}' for c in cols}
    
    if len(cols) >= 2:
        prev_m, curr_m = cols[-2], cols[-1]
        diff_col = f'MoM 증감 ({curr_m} vs {prev_m})'
        rate_col = 'MoM 증감률(%)'
        pivot_df[diff_col] = pivot_df[curr_m] - pivot_df[prev_m]
        pivot_df[rate_col] = ((pivot_df[curr_m] - pivot_df[prev_m]) / pivot_df[prev_m]) * 100
        format_dict[diff_col] = '{:+,.2f}'
        format_dict[rate_col] = '{:+.2f}%'
        
    st.dataframe(pivot_df.style.format(format_dict, na_rep="-"), use_container_width=True)
    
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
        df_db.to_excel(writer, sheet_name='DB_누적전체', index=False)
        pivot_df.to_excel(writer, sheet_name='피벗요약', index=False)
        
    st.download_button(
        label="📥 누적 원소재 마스터 엑셀 다운로드 (백업)",
        data=out_buf.getvalue(),
        file_name=f"원소재_시세_누적마스터_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/octet-stream"
    )
