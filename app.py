import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="글로벌 법인별 원소재 가격 추이 모니터링",
    page_icon="📈",
    layout="wide"
)

# ---------------------------------------------------------
# 1. 원소재 엑셀 데이터 파싱 함수 (원칙 기반 전처리)
# ---------------------------------------------------------
@st.cache_data
def load_and_parse_raw_material_data(file_source):
    excel = pd.ExcelFile(file_source)
    
    # 요약 시트 파싱
    df_raw = pd.read_excel(excel, sheet_name='요약_Summary', header=None)
    
    # 컬럼 정의: 
    # 한국: 7월(col 4), 8월(col 8), 9월(col 12)
    # 중국: 7월(col 16), 8월(col 20), 9월(col 24)
    # 인도: 7월(col 28), 8월(col 32), 9월(col 36)
    subsidiaries = {
        '한국': {'7월': 4, '8월': 8, '9월': 12, 'currency': 'KRW'},
        '중국': {'7월': 16, '8월': 20, '9월': 24, 'currency': 'CNY'},
        '인도': {'7월': 28, '8월': 32, '9월': 36, 'currency': 'INR'}
    }
    
    records = []
    
    # Row 6부터 데이터 시작 (0-index 기준)
    curr_category = ""
    curr_material = ""
    
    for r in range(6, len(df_raw)):
        # Category Forward Fill
        raw_cat = df_raw.iloc[r, 1]
        if pd.notna(raw_cat) and str(raw_cat).strip() != "":
            curr_category = str(raw_cat).split('\n')[0].replace(' / ', '/').strip()
            
        # Material Forward Fill
        raw_mat = df_raw.iloc[r, 2]
        if pd.notna(raw_mat) and str(raw_mat).strip() != "":
            curr_material = str(raw_mat).split('\n')[0].strip()
        elif pd.isna(raw_mat) and curr_category in ['Cu(Kg)', 'LME', 'TTS', '황동원소재/Brass', 'Sn']:
            curr_material = curr_category

        grade = str(df_raw.iloc[r, 3]).strip() if pd.notna(df_raw.iloc[r, 3]) else '-'
        
        # 품목 식별명 생성
        item_name = f"[{curr_category}] {curr_material}" + (f" ({grade})" if grade != '-' else "")
        
        for sub, month_map in subsidiaries.items():
            for m_name, col_idx in month_map.items():
                if m_name == 'currency':
                    continue
                price = df_raw.iloc[r, col_idx]
                curr = df_raw.iloc[r, col_idx + 1]
                supplier = df_raw.iloc[r, col_idx + 2]
                memo = df_raw.iloc[r, col_idx + 3]
                
                # 유효 숫자 검증
                try:
                    price_val = float(price) if pd.notna(price) else np.nan
                except ValueError:
                    price_val = np.nan
                    
                records.append({
                    '법인': sub,
                    '구분': curr_category,
                    '재질': curr_material,
                    'GRADE': grade,
                    '품목명': item_name,
                    '월': m_name,
                    '단가': price_val,
                    '통화': curr if pd.notna(curr) else month_map['currency'],
                    '공급사명': str(supplier).strip() if pd.notna(supplier) else '',
                    'Memo': str(memo).strip() if pd.notna(memo) else ''
                })
                
    df_result = pd.DataFrame(records)
    return df_result

# ---------------------------------------------------------
# 2. 메인 애플리케이션 레이아웃
# ---------------------------------------------------------
st.title("📊 글로벌 법인별 원소재 가격 추이 분석 대시보드")
st.caption("자동차 모터용 주요 비철금속(Cu, Al, 황동), 사출수지 및 LME/환율 추이 통합 모니터링")

# 사이드바: 파일 업로드 및 필터링
st.sidebar.header("📁 데이터 소스 관리")
uploaded_file = st.sidebar.file_uploader("원소재 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df_data = load_and_parse_raw_material_data(uploaded_file)
else:
    default_filename = "[인도법인]법인별 원소재 가격 추이_260930_KR-EN R1.xlsx"
    try:
        df_data = load_and_parse_raw_material_data(default_filename)
        st.sidebar.success(f"기준 마스터 로드 완료: `{default_filename}`")
    except Exception as e:
        st.error(f"파일을 찾을 수 없습니다. 원본 엑셀 파일을 업로드해주세요. 에러: {e}")
        st.stop()

# ---------------------------------------------------------
# 3. 중국 법인 미접수 상태 점검 및 경고 처리
# ---------------------------------------------------------
cn_valid_count = df_data[(df_data['법인'] == '중국') & (df_data['단가'].notna())].shape[0]

if cn_valid_count == 0:
    st.warning("⚠️ **[구매 현황 공지]** 중국 법인의 당월/분기 원소재 단가는 **현재 미접수 상태**입니다. 수신 완료 시 엑셀 업데이트 후 재반영됩니다.")

# 필터링 섹션
st.sidebar.header("🔍 원소재 필터 조건")
selected_subs = st.sidebar.multiselect("대상 법인", options=df_data['법인'].unique(), default=['한국', '인도'])
selected_cats = st.sidebar.multiselect("원소재 구분", options=df_data['구분'].unique(), default=df_data['구분'].unique())

filtered_df = df_data[(df_data['법인'].isin(selected_subs)) & (df_data['구분'].isin(selected_cats))]
items_available = filtered_df['품목명'].unique()
selected_items = st.sidebar.multiselect("상세 품목(재질/GRADE)", options=items_available, default=items_available[:4] if len(items_available) >= 4 else items_available)

final_df = filtered_df[filtered_df['품목명'].isin(selected_items)]

# ---------------------------------------------------------
# 4. KPI 지표 (한국 및 인도 9월 MoM 변동 분석)
# ---------------------------------------------------------
st.subheader("📌 9월 MoM(전월비) 주요 원소재 변동 현황")

kpi_cols = st.columns(4)
cu_kr = df_data[(df_data['법인'] == '한국') & (df_data['구분'] == 'Cu(Kg)')]
cu_in = df_data[(df_data['법인'] == '인도') & (df_data['구분'] == 'Cu(Kg)')]
lme_data = df_data[(df_data['법인'] == '한국') & (df_data['구분'] == 'LME')]
tts_data = df_data[(df_data['법인'] == '한국') & (df_data['구분'] == 'TTS')]

def get_mom_metrics(sub_df):
    p8 = sub_df[sub_df['월'] == '8월']['단가'].values
    p9 = sub_df[sub_df['월'] == '9월']['단가'].values
    if len(p8) > 0 and len(p9) > 0 and pd.notna(p8[0]) and pd.notna(p9[0]):
        diff = p9[0] - p8[0]
        pct = (diff / p8[0]) * 100
        return p9[0], diff, pct
    return None, None, None

# Card 1: 한국 Cu
p9, diff, pct = get_mom_metrics(cu_kr)
if p9:
    kpi_cols[0].metric("한국 Cu(Kg)", f"{p9:,.0f} KRW", f"{diff:+,.0f} ({pct:+.2f}%)")

# Card 2: 인도 Cu (INR)
p9, diff, pct = get_mom_metrics(cu_in)
if p9:
    kpi_cols[1].metric("인도 Cu(Kg)", f"{p9:,.2f} INR", f"{diff:+,.2f} ({pct:+.2f}%)")

# Card 3: LME
p9, diff, pct = get_mom_metrics(lme_data)
if p9:
    kpi_cols[2].metric("LME 동지표", f"${p9:,.2f}", f"{diff:+,.2f} ({pct:+.2f}%)")

# Card 4: 환율 TTS
p9, diff, pct = get_mom_metrics(tts_data)
if p9:
    kpi_cols[3].metric("원/달러 TTS", f"{p9:,.2f} KRW", f"{diff:+,.2f} ({pct:+.2f}%)")

st.markdown("---")

# ---------------------------------------------------------
# 5. 시계열 단가 추이 차트 (Plotly Interactive)
# ---------------------------------------------------------
st.subheader("📈 품목별 월별 단가 추이")

if not final_df.empty:
    fig = px.line(
        final_df.dropna(subset=['단가']),
        x='월',
        y='단가',
        color='품목명',
        line_dash='법인',
        markers=True,
        hover_data=['통화', '공급사명'],
        title="법인별·원소재별 가격 변동 추이 (7월 ~ 9월)"
    )
    fig.update_layout(
        hovermode="x unified",
        xaxis_title="기준월",
        yaxis_title="단가 (각 법인 통화 기준)",
        legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="center", x=0.5)
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("선택된 필터 조건에 부합하는 데이터가 없습니다.")

# ---------------------------------------------------------
# 6. 원가 검증 및 데이터 테이블 (Pivot Table)
# ---------------------------------------------------------
st.subheader("📋 원소재 단가 세부 내역표 (Summary View)")

# Pivot Table 생성
pivot_df = final_df.pivot_table(
    index=['법인', '구분', '재질', 'GRADE', '통화', '공급사명'],
    columns='월',
    values='단가',
    aggfunc='first'
).reset_index()

# 7월, 8월, 9월 컬럼 순서 정렬 및 MoM 계산
for col in ['7월', '8월', '9월']:
    if col not in pivot_df.columns:
        pivot_df[col] = np.nan

pivot_df = pivot_df[['법인', '구분', '재질', 'GRADE', '통화', '공급사명', '7월', '8월', '9월']]
pivot_df['MoM 증감률(%)'] = ((pivot_df['9월'] - pivot_df['8월']) / pivot_df['8월']) * 100

st.dataframe(
    pivot_df.style.format({
        '7월': '{:,.2f}',
        '8월': '{:,.2f}',
        '9월': '{:,.2f}',
        'MoM 증감률(%)': '{:+.2f}%'
    }, na_rep="-"),
    use_container_width=True
)

# 엑셀 다운로드 기능
output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    pivot_df.to_excel(writer, sheet_name='Price_Trend_Analysis', index=False)
st.download_button(
    label="📥 정제 분석 데이터 엑셀 다운로드",
    data=output.getvalue(),
    file_name="원소재_가격추이_분석_정리.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
