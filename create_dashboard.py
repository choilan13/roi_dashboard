"""
create_dashboard.py
전체 가공 시트 → ROI 피봇 대시보드 HTML 생성
"""

from __future__ import annotations
import sys
from pathlib import Path
import numpy as np
import pandas as pd

BASE           = Path(__file__).parent
CREDENTIALS    = BASE / 'credentials.json'
SPREADSHEET_ID = '17_oeV41FVchyFYUl6dHQyEk689tkzrA840WfWApFkQY'
OUTPUT_SHEET   = '전체 가공'
OUTPUT_HTML    = BASE / 'roi_dashboard.html'

SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

COLOR_BLUE   = '#2563EB'
COLOR_GREEN  = '#059669'
COLOR_RED    = '#DC2626'
COLOR_AMBER  = '#D97706'
COLOR_NAVY   = '#1E3A5F'


# ── 데이터 로드 ────────────────────────────────────────────────────────────────

def load_data() -> pd.DataFrame:
    import gspread
    from google.oauth2.service_account import Credentials
    print('Google Sheets 로드 중...')
    creds = Credentials.from_service_account_file(str(CREDENTIALS), scopes=SCOPES)
    gc    = gspread.authorize(creds)
    ws    = gc.open_by_key(SPREADSHEET_ID).worksheet(OUTPUT_SHEET)
    data  = ws.get_all_values()
    headers, rows = data[0], data[1:]
    df = pd.DataFrame(rows, columns=headers)
    print(f'  {len(df):,}행 로드 완료')
    return df


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    for col in ['매출', '공헌이익', '원가', '수수료', '실제 부과 배송']:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(',', '').str.strip(),
                errors='coerce'
            ).fillna(0)

    df['공헌이익율'] = np.where(
        df['매출'] != 0, df['공헌이익'] / df['매출'], np.nan
    )

    def safe_int(s, pat):
        try:
            return int(str(s).replace(pat, ''))
        except Exception:
            return 99
    df['월_n']  = df['월'].apply(lambda x: safe_int(x, '월'))
    df['주차_n'] = df['주차'].apply(lambda x: safe_int(x, '주'))
    df['월주']   = df['월'] + ' ' + df['주차']
    df['월주_n'] = df['월_n'] * 10 + df['주차_n']

    # 빈 SKU명 → "(미분류)"
    df['SKU명'] = df['SKU명'].replace('', '(미분류)').fillna('(미분류)')
    df['브랜드'] = df['브랜드'].replace('', '(미분류)').fillna('(미분류)')
    df['채널명'] = df['채널명'].replace('', '(미분류)').fillna('(미분류)')
    return df


# ── 포맷 헬퍼 ──────────────────────────────────────────────────────────────────

def krw(v):
    if pd.isna(v): return '-'
    return f'₩{int(v):,}'

def pct(v):
    if pd.isna(v): return '-'
    return f'{v:.1%}'

def pc(v):
    return COLOR_GREEN if (not pd.isna(v) and v >= 0) else COLOR_RED


# ── 피봇 계산 ──────────────────────────────────────────────────────────────────

def pivot(df, group_col):
    p = df.groupby(group_col, as_index=False).agg(
        매출=('매출', 'sum'), 공헌이익=('공헌이익', 'sum')
    )
    p['공헌이익율'] = np.where(p['매출'] != 0, p['공헌이익'] / p['매출'], np.nan)
    return p.sort_values('매출', ascending=False).reset_index(drop=True)


def pivot_month(df):
    p = df.groupby(['월_n', '월'], as_index=False).agg(
        매출=('매출', 'sum'), 공헌이익=('공헌이익', 'sum')
    ).sort_values('월_n')
    p['공헌이익율'] = np.where(p['매출'] != 0, p['공헌이익'] / p['매출'], np.nan)
    return p


def pivot_week(df):
    p = df.groupby(['월주_n', '월주'], as_index=False).agg(
        매출=('매출', 'sum'), 공헌이익=('공헌이익', 'sum')
    ).sort_values('월주_n')
    p['공헌이익율'] = np.where(p['매출'] != 0, p['공헌이익'] / p['매출'], np.nan)
    return p


# ── 피봇 테이블 Figure ─────────────────────────────────────────────────────────

def table_fig(p, label_col, title):
    import plotly.graph_objects as go
    n = len(p)
    contrib_colors = [pc(v) for v in p['공헌이익']]
    rate_colors    = [pc(v) for v in p['공헌이익율']]
    row_bg = ['#F8FAFC' if i % 2 == 0 else 'white' for i in range(n)]

    fig = go.Figure(data=[go.Table(
        columnwidth=[3, 2, 2, 1.5],
        header=dict(
            values=[f'<b>{label_col}</b>', '<b>매출</b>', '<b>공헌이익</b>', '<b>공헌이익율</b>'],
            fill_color=COLOR_NAVY,
            font=dict(color='white', size=12),
            align=['left', 'right', 'right', 'center'],
            height=34,
        ),
        cells=dict(
            values=[
                p[label_col],
                [krw(v) for v in p['매출']],
                [krw(v) for v in p['공헌이익']],
                [pct(v) for v in p['공헌이익율']],
            ],
            fill_color=[row_bg, row_bg, contrib_colors, rate_colors],
            font=dict(color=['#1E293B', '#1E293B', 'white', 'white'], size=11),
            align=['left', 'right', 'right', 'center'],
            height=28,
        ),
    )])
    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color=COLOR_NAVY)),
        height=min(max(200, n * 30 + 100), 800),
        margin=dict(l=0, r=0, t=40, b=0),
        template='plotly_white',
    )
    return fig


# ── 차트 Figure ────────────────────────────────────────────────────────────────

def trend_fig(mo, title):
    import plotly.graph_objects as go
    fig = go.Figure()
    fig.add_bar(
        x=mo['월'], y=mo['매출'], name='매출',
        marker_color=COLOR_BLUE, opacity=0.85, offsetgroup=0,
    )
    fig.add_bar(
        x=mo['월'], y=mo['공헌이익'], name='공헌이익',
        marker_color=[pc(v) for v in mo['공헌이익']], opacity=0.85, offsetgroup=1,
    )
    fig.add_scatter(
        x=mo['월'], y=mo['공헌이익율'], name='공헌이익율',
        mode='lines+markers+text',
        text=[pct(v) for v in mo['공헌이익율']],
        textposition='top center', textfont=dict(size=11, color=COLOR_AMBER),
        yaxis='y2', line=dict(color=COLOR_AMBER, width=2.5), marker=dict(size=9),
    )
    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color=COLOR_NAVY)),
        barmode='group',
        yaxis=dict(title='금액 (원)', tickformat=',d'),
        yaxis2=dict(title='공헌이익율', overlaying='y', side='right',
                    tickformat='.0%', showgrid=False),
        legend=dict(orientation='h', y=1.12),
        height=430, template='plotly_white',
        margin=dict(t=60),
    )
    return fig


def week_trend_fig(wk):
    import plotly.graph_objects as go
    bar_colors = [COLOR_BLUE] * len(wk)
    fig = go.Figure()
    fig.add_bar(x=wk['월주'], y=wk['매출'], name='매출',
                marker_color=bar_colors, opacity=0.8)
    fig.add_scatter(
        x=wk['월주'], y=wk['공헌이익율'], name='공헌이익율',
        mode='lines+markers', yaxis='y2',
        line=dict(color=COLOR_AMBER, width=2), marker=dict(size=7),
    )
    fig.update_layout(
        title=dict(text='주차별 매출 / 공헌이익율', font=dict(size=14, color=COLOR_NAVY)),
        yaxis=dict(title='매출 (원)', tickformat=',d'),
        yaxis2=dict(title='공헌이익율', overlaying='y', side='right',
                    tickformat='.0%', showgrid=False),
        legend=dict(orientation='h', y=1.12),
        height=400, template='plotly_white',
        xaxis=dict(tickangle=-45),
        margin=dict(t=60, b=80),
    )
    return fig


def hbar_fig(p, label_col, title, top_n=None):
    import plotly.graph_objects as go
    d = p.head(top_n) if top_n else p
    d = d.sort_values('매출')
    fig = go.Figure()
    fig.add_bar(y=d[label_col], x=d['매출'], name='매출',
                orientation='h', marker_color=COLOR_BLUE, opacity=0.85)
    fig.add_bar(y=d[label_col], x=d['공헌이익'], name='공헌이익',
                orientation='h',
                marker_color=[pc(v) for v in d['공헌이익']], opacity=0.85)
    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color=COLOR_NAVY)),
        barmode='group', height=max(380, len(d) * 38 + 120),
        xaxis=dict(tickformat=',d'),
        legend=dict(orientation='h', y=1.08),
        template='plotly_white',
        margin=dict(t=60),
    )
    return fig


def heatmap_fig(df):
    import plotly.graph_objects as go
    cross = df.pivot_table(
        values='매출', index='채널명', columns='월',
        aggfunc='sum', fill_value=0,
    )
    sorted_cols = sorted(cross.columns, key=lambda x: int(x.replace('월', '')))
    cross = cross[sorted_cols]
    sorted_rows = cross.sum(axis=1).sort_values(ascending=False).index
    cross = cross.loc[sorted_rows]

    fig = go.Figure(data=go.Heatmap(
        z=cross.values,
        x=cross.columns.tolist(),
        y=cross.index.tolist(),
        colorscale='Blues',
        text=[[krw(v) for v in row] for row in cross.values],
        texttemplate='%{text}',
        textfont=dict(size=10),
        hovertemplate='%{y} / %{x}<br>매출: %{text}<extra></extra>',
    ))
    fig.update_layout(
        title=dict(text='채널 × 월 매출 히트맵', font=dict(size=14, color=COLOR_NAVY)),
        height=max(400, len(cross) * 34 + 120),
        template='plotly_white',
        margin=dict(t=60),
    )
    return fig


# ── 대시보드 HTML 생성 ─────────────────────────────────────────────────────────

def build_dashboard(df: pd.DataFrame) -> Path:
    import plotly.io as pio

    mo_pv  = pivot_month(df)
    wk_pv  = pivot_week(df)
    ch_pv  = pivot(df, '채널명')
    br_pv  = pivot(df, '브랜드')
    sku_pv = pivot(df, 'SKU명')

    total_sales   = df['매출'].sum()
    total_contrib = df['공헌이익'].sum()
    total_rate    = total_contrib / total_sales if total_sales else 0

    # 차트
    charts = [
        trend_fig(mo_pv, '월별 매출 / 공헌이익 추이'),
        week_trend_fig(wk_pv),
        hbar_fig(ch_pv, '채널명', '채널별 매출 / 공헌이익'),
        hbar_fig(br_pv, '브랜드', '브랜드별 매출 / 공헌이익'),
        hbar_fig(sku_pv, 'SKU명', 'SKU별 매출 / 공헌이익 (상위 20)', top_n=20),
        heatmap_fig(df),
    ]

    # 테이블
    tables = [
        table_fig(mo_pv[['월', '매출', '공헌이익', '공헌이익율']], '월', '월별 피봇'),
        table_fig(wk_pv[['월주', '매출', '공헌이익', '공헌이익율']], '월주', '주차별 피봇'),
        table_fig(ch_pv, '채널명', '채널별 피봇'),
        table_fig(br_pv, '브랜드', '브랜드별 피봇'),
        table_fig(sku_pv.head(60), 'SKU명', 'SKU별 피봇 (상위 60)'),
    ]

    def to_html(fig):
        return pio.to_html(fig, include_plotlyjs=False, full_html=False)

    charts_html = '\n'.join(f'<div class="section">{to_html(f)}</div>' for f in charts)
    tables_html = '\n'.join(f'<div class="section">{to_html(f)}</div>' for f in tables)

    def kpi_card(label, value, color=''):
        style = f'style="color:{color}"' if color else ''
        return f'''<div class="kpi"><div class="kpi-label">{label}</div>
<div class="kpi-value" {style}>{value}</div></div>'''

    kpi_row = (
        kpi_card('총 매출', krw(total_sales))
        + kpi_card('총 공헌이익', krw(total_contrib),
                   COLOR_GREEN if total_contrib >= 0 else COLOR_RED)
        + kpi_card('공헌이익율', pct(total_rate),
                   COLOR_GREEN if total_rate >= 0 else COLOR_RED)
        + kpi_card('채널 수', f'{df["채널명"].nunique()}개')
        + kpi_card('브랜드 수', f'{df["브랜드"].nunique()}개')
        + kpi_card('SKU 수', f'{df["SKU명"].nunique()}개')
        + kpi_card('기간', f'{df["월"].min()} ~ {df["월"].max()}')
    )

    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ROI 대시보드 2026</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Malgun Gothic","Segoe UI",sans-serif;background:#F0F4F8;color:#1E293B}}
header{{background:{COLOR_NAVY};color:white;padding:22px 32px;display:flex;align-items:baseline;gap:16px}}
header h1{{font-size:1.5rem;font-weight:700}}
header span{{font-size:.85rem;opacity:.65}}
.kpi-row{{display:flex;gap:12px;padding:20px 32px;flex-wrap:wrap}}
.kpi{{background:white;border-radius:10px;padding:16px 22px;flex:1;min-width:140px;
       box-shadow:0 1px 3px rgba(0,0,0,.08)}}
.kpi-label{{font-size:.72rem;color:#64748B;text-transform:uppercase;letter-spacing:.06em}}
.kpi-value{{font-size:1.5rem;font-weight:700;margin-top:5px;color:{COLOR_NAVY}}}
.tabs{{display:flex;gap:6px;padding:0 32px 16px;border-bottom:1px solid #E2E8F0;margin-bottom:0}}
.tab{{padding:8px 20px;border-radius:8px 8px 0 0;cursor:pointer;font-size:.9rem;font-weight:500;
      background:white;border:1px solid #E2E8F0;border-bottom:none;color:#64748B;transition:.15s}}
.tab:hover{{background:#EFF6FF;color:{COLOR_BLUE}}}
.tab.active{{background:{COLOR_NAVY};color:white;border-color:{COLOR_NAVY}}}
.main{{padding:20px 32px;display:flex;flex-direction:column;gap:16px}}
.section{{background:white;border-radius:10px;padding:16px;
          box-shadow:0 1px 3px rgba(0,0,0,.08);overflow:hidden}}
.pane{{display:none}}.pane.active{{display:flex;flex-direction:column;gap:16px}}
</style>
</head>
<body>
<header>
  <h1>📊 ROI 대시보드 2026</h1>
  <span>전체 가공 기준 · {len(df):,}건 · 1월 ~ 5월 3주</span>
</header>
<div class="kpi-row">{kpi_row}</div>
<div class="tabs" style="padding-top:16px">
  <div class="tab active" onclick="show('charts',this)">📈 차트</div>
  <div class="tab" onclick="show('tables',this)">📋 피봇 테이블</div>
</div>
<div class="main">
  <div id="charts" class="pane active">{charts_html}</div>
  <div id="tables" class="pane">{tables_html}</div>
</div>
<script>
function show(id,el){{
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  el.classList.add('active');
}}
</script>
</body>
</html>'''

    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print(f'저장: {OUTPUT_HTML}')
    return OUTPUT_HTML


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    df  = load_data()
    df  = preprocess(df)
    out = build_dashboard(df)
    import webbrowser
    webbrowser.open(out.as_uri())
    print('브라우저 오픈')


if __name__ == '__main__':
    main()
