import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
from datetime import datetime
from io import BytesIO
import os

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="글로벌 법인별 원소재 시세 Trend 누적 모니터링",
    page_icon="📈",
    layout="wide"
)

DB_PATH = "raw_materials_trend.db"

# ---------------------------------------------------------
# 1. SQLite Database Initialization & Storage Engine
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

def save_to_db(df_records):
    if df_records.empty:
        return 0
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    count = 0
    for _, row in df_records.iterrows():
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
            row['subsidiary'], row['category'], row['material'], row['grade'],
            row['month'], row['price'], row['currency'], row['supplier'], row['memo'], now_str
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
# 2. Excel Parsing Engine (요약 시트 + HS India Grades 시트)
# ---------------------------------------------------------
def parse_raw_material_excel(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    records = []
    
    # [A] 요약_Summary 시트 파싱
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
                    except (ValueError, TypeError):
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

    # [B] HS India Grades 시트 (인도 로컬 사출 엔프라 수지)
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
                except (ValueError, TypeError):
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
# 3. Main State Initialization & Auto-Seeding
# ---------------------------------------------------------
init_db()

# 사이드바: 모바일 업로드
st.sidebar.markdown("### 📱 모바일 데이터 업로드 & 누적")
st.sidebar.info("사내망 보안 차단 시, 모바일로 접속해 엑셀을 업로드하면 DB에 자동 누적·갱신됩니다.")

uploaded_file = st.sidebar.file_uploader("최신 원소재 엑셀 파일 (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    with st.spinner("엑셀 데이터 파싱 및 SQLite DB 누적 중..."):
        parsed_df = parse_raw_material_excel(uploaded_file)
        saved_cnt = save_to_db(parsed_df)
        st.sidebar.success(f"총 {saved_cnt}건 누적 저장 완료!")

# DB에서 누적 데이터 로드
df_db = load_data_from_db()

# DB가 비어있는 경우 로컬 기준 엑셀 자동 시드 (최초 1회)
if df_db.empty:
    for default_name in ["[인도법인]법인별 원소재 가격 추이_260930_KR-EN R1.xlsx", "raw_material_sample.xlsx"]:
        if os.path.exists(default_name):
            parsed_init = parse_raw_material_excel(default_name)
            save_to_db(parsed_init)
            df_db = load_data_from_db()
            break

# ---------------------------------------------------------
# 4. 사이드바 필터 (재질별 드롭다운)
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 시세 조회 조건 설정")

if not df_db.empty:
    # 1) 재질 드롭다운 (Selectbox)
    available_materials = sorted([m for m in df_db['material'].dropna().unique() if m != ''])
    selected_material = st.sidebar.selectbox(
        "📌 재질(Material) 선택",
        options=["전체 재질"] + available_materials,
        index=0
    )
    
    # 2) 대상 법인 다중 선택
    available_subs = list(df_db['subsidiary'].unique())
    selected_subs = st.sidebar.multiselect(
        "🏢 대상 법인 선택",
        options=available_subs,
        default=[s for s in ['한국', '인도'] if s in available_subs]
    )
    
    # 필터링 적용
    mask = df_db['subsidiary'].isin(selected_subs)
    if selected_material != "전체 재질":
        mask = mask & (df_db['material'] == selected_material)
    sub_filtered_df = df_db[mask]
    
    # 3) 세부 GRADE 필터
    available_grades = sorted(sub_filtered_df['grade'].unique())
    selected_grades = st.sidebar.multiselect(
        "🏷️ GRADE 선택",
        options=available_grades,
        default=available_grades
    )
    
    final_df = sub_filtered_df[sub_filtered_df['grade'].isin(selected_grades)].copy()
else:
    final_df = pd.DataFrame()

# ---------------------------------------------------------
# 5. 메인 대시보드 화면
# ---------------------------------------------------------
st.title("📊 글로벌 법인별 원소재 단가 Trend 분석")
st.caption("소형 모터 주요 원자재(Cu, Al, 황동, 사출수지) 시계열 누적 모니터링")

# 중국 법인 미접수 안내 배너
if not df_db.empty:
    cn_prices = df_db[(df_db['subsidiary'] == '중국') & (df_db['price'].notna())]
    if cn_prices.empty:
        st.warning("⚠️ **[구매 현황 공지] 중국 법인 원소재 단가는 현재 미접수 상태입니다.** 접수 후 모바일 업로드 시 즉시 통합 반영됩니다.")

# ---------------------------------------------------------
# 6. Trend 꺾은선 그래프 (Line Chart)
# ---------------------------------------------------------
st.subheader(f"📈 {'[' + selected_material + ']' if selected_material != '전체 재질' else '원소재'} 단가 변동 Trend")

valid_plot_df = final_df.dropna(subset=['price']).copy()

if not valid_plot_df.empty:
    # 월 순서 정렬
    month_order = {'7월': 1, '8월': 2, '9월': 3, '10월': 4, '11월': 5, '12월': 6}
    valid_plot_df['month_rank'] = valid_plot_df['month'].map(lambda x: month_order.get(x, 99))
    valid_plot_df = valid_plot_df.sort_values(by=['month_rank', 'subsidiary', 'material', 'grade'])
    
    # 선 식별 라벨
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
        title=f"월별 단가 추이 모니터링 ({selected_material})"
    )
    
    fig.update_layout(
        xaxis_title="기준월 (Month)",
        yaxis_title="단가 (각 법인 결제 통화 기준)",
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
    st.info("선택된 필터 조건에 표시할 유효 단가 데이터가 없습니다.")

st.markdown("---")

# ---------------------------------------------------------
# 7. 세부 피벗 테이블 및 MoM(전월비) 변동률 산출
# ---------------------------------------------------------
st.subheader("📋 세부 단가 피벗 테이블 및 MoM(전월비) 분석")

if not final_df.empty:
    pivot_df = final_df.pivot_table(
        index=['subsidiary', 'category', 'material', 'grade', 'currency', 'supplier'],
        columns='month',
        values='price',
        aggfunc='first'
    ).reset_index()
    
    cols = [c for c in ['7월', '8월', '9월'] if c in pivot_df.columns]
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
    
    # 엑셀 백업 다운로드
    out_buf = BytesIO()
    with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
        df_db.to_excel(writer, sheet_name='DB_전체누적데이터', index=False)
        pivot_df.to_excel(writer, sheet_name='현재조회_피벗요약', index=False)
        
    st.download_button(
        label="📥 누적 원소재 마스터 DB 엑셀 다운로드 (백업)",
        data=out_buf.getvalue(),
        file_name=f"원소재_시세_누적마스터_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd
