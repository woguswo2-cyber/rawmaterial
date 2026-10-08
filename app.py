import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
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
        p_val = float(row['price']) if pd.notna(row['price']) else None
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
    df = pd.read_sql_query("SELECT * FROM material_prices ORDER BY category, material, grade, month", conn)
    conn.close()
    return df

# ---------------------------------------------------------
# 2. 엑셀 복사/붙여넣기 파싱 엔진 (Ctrl+C -> Ctrl+V)
# ---------------------------------------------------------
def parse_clipboard_text(pasted_text, sub_default="한국", curr_default="KRW"):
    lines = [line for line in pasted_text.strip().splitlines() if line.strip()]
    if not lines:
        return pd.DataFrame()
    
    # 탭(\t) 또는 쉼표(,) 구분자 감지
    sep = '\t' if '\t' in lines[0] else ','
    raw_df = pd.read_csv(io.StringIO(pasted_text), sep=sep)
    raw_df.columns = [str(c).strip() for c in raw_df.columns]
    
    # 시세 월 컬럼 vs 기본 정보 컬럼 자동 분류
    month_cols = []
    meta_cols = []
    
    for c in raw_df.columns:
        if any(m in c for m in ['월', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', '202']):
            month_cols.append(c)
        elif c in ['법인', '구분', '재질', 'GRADE', 'Grade', '품목', '품목명', 'Material', '공급사', '공급사명', '통화', '결제통화', '비고', 'Memo']:
            meta_cols.append(c)
        else:
            # 수치 데이터 비율이 높으면 시세 월 컬럼으로 자동 판정
            num_ratio = pd.to_numeric(raw_df[c].astype(str).str.replace(',', '').str.replace(' ', ''), errors='coerce').notna().mean()
            if num_ratio > 0.4:
                month_cols.append(c)
            else:
                meta_cols.append(c)
                
    records = []
    for _, row in raw_df.iterrows():
        # 재질 및 구분 감지
        mat = str(row.get('재질', row.get('품목', row.get('구분', row.get('Material', '기타소재'))))).strip()
        cat = str(row.get('구분', mat)).strip()
        grade = str(row.get('GRADE', row.get('Grade', 'STD'))).strip()
        if grade in ['nan', '', 'None']:
            grade = 'STD'
        sub = str(row.get('법인', sub_default)).strip()
        curr = str(row.get('통화', row.get('결제통화', curr_default))).strip()
        supp = str(row.get('공급사', row.get('공급사명', ''))).strip()
        memo = str(row.get('비고', row.get('Memo', ''))).strip()
        
        for m in month_cols:
            val = row[m]
            try:
                p_val = float(str(val).replace(',', '').strip())
            except:
                p_val = np.nan
                
            records.append({
                'subsidiary': sub,
                'category': cat,
                'material': mat,
                'grade': grade,
                'month': m,
                'price': p_val,
                'currency': curr,
                'supplier': supp if supp != 'nan' else '',
                'memo': memo if memo != 'nan' else ''
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
                    except:
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
                except:
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
# 4. App Execution & DB Initial Setup
# ---------------------------------------------------------
init_db()

# DB가 비어있는 경우 기본 엑셀이 있으면 자동 시드
df_db = load_data_from_db()
if df_db.empty:
    for default_name in ["[인도법인]법인별 원소재 가격 추이_260930_KR-EN R1.xlsx", "raw_material_sample.xlsx"]:
        if os.path.exists(default_name):
            init_records = parse_raw_material_excel(default_name)
            save_records_to_db(init_records)
            df_db = load_data_from_db()
            break

# ---------------------------------------------------------
# 5. 데이터 입력 탭 (엑셀 복사/붙여넣기 vs 모바일 파일 업로드)
# ---------------------------------------------------------
st.title("📊 글로벌 법인별 원소재 단가 Trend 분석 대시보드")
st.caption("사출수지, 비철금속(Cu, Al, 황동) 월별 시세 추이 모니터링 & DB 누적 관리")

input_tab1, input_tab2 = st.tabs(["📋 엑셀 데이터 직접 붙여넣기 (사내망/보안 환경용)", "📁 엑셀 파일 업로드 (모바일/기타)"])

with input_tab1:
    st.info("💡 **사용법**: 사내 엑셀에서 단가 표 영역을 드래그하여 **복사(Ctrl+C)**한 후, 아래 입력창에 **붙여넣기(Ctrl+V)**하세요.")
    p_col1, p_col2 = st.columns(2)
    with p_col1:
        paste_sub = st.selectbox("기본 법인 지정", ["한국", "인도", "중국"], index=0)
    with p_col2:
        paste_curr = st.selectbox("결제 통화 지정", ["KRW", "INR", "USD", "CNY"], index=0)
        
    pasted_text = st.text_area(
        "엑셀 표 붙여넣기 창",
        placeholder="예시:\n재질\tGRADE\t7월\t8월\t9월\nPP-TD20\tJF1512\t2015\t1765\t1785\nALDC12\tSTD\t5376\t5148\t4777",
        height=150
    )
    
    if st.button("🚀 붙여넣은 데이터로 즉시 차트 반영 & DB 누적 저장"):
        if pasted_text.strip():
            parsed_paste_df = parse_clipboard_text(pasted_text, sub_default=paste_sub, curr_default=paste_curr)
            if not parsed_paste_df.empty:
                saved = save_records_to_db(parsed_paste_df)
                st.success(f"성공! 총 {saved}건의 시세 데이터가 누적 DB에 저장되었습니다.")
                df_db = load_data_from_db()
            else:
                st.error("데이터 파싱에 실패했습니다. 엑셀의 헤더(재질, 월 등)를 포함하여 복사했는지 확인해주세요.")
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

if not df_db.empty:
    # 1) 재질 드롭다운
    material_list = sorted([m for m in df_db['material'].dropna().unique() if m != ''])
    selected_material = st.sidebar.selectbox(
        "📌 재질(Material) 선택",
        options=["전체 재질"] + material_list,
        index=0
    )
    
    # 2) 법인 필터
    subs_list = list(df_db['subsidiary'].unique())
    selected_subs = st.sidebar.multiselect(
        "🏢 대상 법인 선택",
        options=subs_list,
        default=[s for s in ['한국', '인도'] if s in subs_list]
    )
    
    # 1차 필터링
    mask = df_db['subsidiary'].isin(selected_subs)
    if selected_material != "전체 재질":
        mask = mask & (df_db['material'] == selected_material)
    sub_filtered = df_db[mask]
    
    # 3) 세부 GRADE 필터
    grade_list = sorted(sub_filtered['grade'].unique())
    selected_grades = st.sidebar.multiselect(
        "🏷️ GRADE 선택",
        options=grade_list,
        default=grade_list
    )
    
    final_df = sub_filtered[sub_filtered['grade'].isin(selected_grades)].copy()
else:
    final_df = pd.DataFrame()

st.markdown("---")

# ---------------------------------------------------------
# 7. Trend 꺾은선 그래프
# ---------------------------------------------------------
st.subheader(f"📈 {'[' + selected_material + ']' if selected_material != '전체 재질' else '주요 원소재'} 단가 변동 Trend (꺾은선 그래프)")

valid_plot_df = final_df.dropna(subset=['price']).copy()

if not valid_plot_df.empty:
    # 월 순서 정렬 매핑
    month_order = {'7월': 1, '8월': 2, '9월': 3, '10월': 4, '11월': 5, '12월': 6}
    valid_plot_df['month_rank'] = valid_plot_df['month'].map(lambda x: month_order.get(x, 99))
    valid_plot_df = valid_plot_df.sort_values(by=['month_rank', 'subsidiary', 'material', 'grade'])
    
    # 라벨 포맷
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
    st.info("조회할 유효 단가 데이터가 없습니다. 상단에서 엑셀 데이터를 붙여넣거나 필터를 조정하세요.")

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
    
    cols = [c for c in ['7월', '8월', '9월', '10월'] if c in pivot_df.columns]
    if len(cols) >= 2:
        prev_m, curr_m = cols[-2], cols[-1]
        pivot_df[f'MoM 증감 ({curr_m} vs {prev_m})'] = pivot_df[curr_m] - pivot_df[prev_m]
        pivot_df['MoM 증감률(%)'] = ((pivot_df[curr_m] - pivot_df[prev_m]) / pivot_df[prev_m]) * 100
        
    st.dataframe(
        pivot_df.style.format({
            c: '{:,.2f}' for c in cols
        } | {
            f'MoM 증감 ({curr_m} vs {prev_m})': '{:+,.2f}',
            'MoM 증감률(%)': '{:+.2f}%'
        }, na_rep="-"),
        use_container_width=True
    )
    
    # 안전한 엑셀 백업 다운로드
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
