"""
수학본부 성과지표 PPT 자동 생성 스크립트
- 템플릿 PPTX를 복사하고 계산된 지표 데이터로 테이블을 채움
- 슬라이드3 차트: 최근 4주차 재수강률/신규전액환불률 라인차트 이미지 생성
- 슬라이드18 매출액: 빈칸 처리
- 영어 데이터: 빈칸 처리
"""

import copy
import json
import sys
import os
import re
import shutil
import io
from lxml import etree
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# ─── 네임스페이스 ───────────────────────────────
NS = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
for prefix, uri in NS.items():
    etree.register_namespace(prefix, uri)

PPTX_TEMPLATE = '/mnt/user-data/uploads/_최종__26년_4월_2주차_수학온택트_성과지표보고.pptx'
UNPACKED_DIR  = '/home/claude/unpacked_gen'
OUTPUT_PPTX   = '/home/claude/output_report.pptx'

# ─── 실장 고정 매핑 ─────────────────────────────
DEPT_MANAGER = {
    '초등팀': '박재국', '초등실': '박재국',
    '중등1A팀': '박재국', '중등1실': '박재국',
    '중등2A팀': '박재국', '중등2실': '박재국',
    '중등3A팀': '조홍래', '중등3실': '조홍래',
    '고등1A팀': '조홍래', '고등1실': '조홍래',
    '고등2A팀': '이상준', '고등2실': '이상준',
    '고등3A팀': '이상준', '고등3실': '이상준',
    'B2B사업팀': '이상준', 'B2B사업실': '이상준',
    '주말팀': '조홍래',
}

# 실 표시 순서 (슬라이드4~9)
DEPT_ORDER_SILL = ['초등실','중등1실','중등2실','중등3실','고등1실','고등2실','고등3실','B2B사업실','주말팀']
# 팀 표시 순서 (슬라이드10~15)
DEPT_ORDER_TEAM = ['초등팀','중등1A팀','중등2A팀','중등3A팀','고등1A팀','고등2A팀','고등3A팀','B2B사업팀','주말팀']

def sill_to_team(sill_name):
    """실 이름 → 팀 이름 변환"""
    return sill_name.replace('실','팀').replace('주말팀','주말팀')

def team_to_sill(team_name):
    return team_name.replace('A팀','실').replace('팀','실')

# ─── XML 헬퍼 ───────────────────────────────────
def get_cell_text(tc, a_ns):
    parts = []
    for t in tc.findall(f'.//{{{a_ns}}}t'):
        if t.text: parts.append(t.text)
    return ''.join(parts)

def set_cell_text(tc, text, a_ns):
    """테이블 셀 텍스트 교체 (첫 번째 run만 수정, 나머지 제거)"""
    # 모든 paragraph 수집
    paras = tc.findall(f'{{{a_ns}}}txBody/{{{a_ns}}}p')
    if not paras:
        return
    # 첫 para 사용, 나머지 제거
    para = paras[0]
    for extra in paras[1:]:
        tc.find(f'{{{a_ns}}}txBody').remove(extra)
    # 해당 para의 run들
    runs = para.findall(f'{{{a_ns}}}r')
    if runs:
        # 첫 run 텍스트 교체
        t_el = runs[0].find(f'{{{a_ns}}}t')
        if t_el is not None:
            t_el.text = str(text)
        # 나머지 run 제거
        for r in runs[1:]:
            para.remove(r)
    else:
        # run 없으면 생성
        r_el = etree.SubElement(para, f'{{{a_ns}}}r')
        rpr = etree.SubElement(r_el, f'{{{a_ns}}}rPr')
        rpr.set('lang', 'ko-KR')
        t_el = etree.SubElement(r_el, f'{{{a_ns}}}t')
        t_el.text = str(text)

def parse_slide(slide_path):
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(slide_path, parser)
    return tree

def save_slide(tree, slide_path):
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8', pretty_print=False)

def get_tables(tree):
    a_ns = NS['a']
    return tree.findall(f'.//{{{a_ns}}}tbl')

# ─── 데이터 받아오기 ────────────────────────────
def load_data(data_json_path):
    with open(data_json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

# ─── 차트 이미지 생성 ───────────────────────────
def create_trend_chart(week_labels, renew_rates, refund_rates, output_path):
    """최근 4주차 재수강률/신규전액환불률 라인 차트 생성"""
    # 한국어 폰트 탐색
    font_candidates = [
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    ]
    font_path = None
    for fp in font_candidates:
        if os.path.exists(fp):
            font_path = fp
            break

    fig, ax = plt.subplots(figsize=(10, 6), facecolor='white')

    x = range(len(week_labels))

    # 재수강률 (파란색)
    color_renew  = '#00BCD4'
    color_refund = '#E53935'

    ax.plot(x, renew_rates,  'o-', color=color_renew,  linewidth=2.5, markersize=8, label='재수강률(%)')
    ax.plot(x, refund_rates, 'o-', color=color_refund, linewidth=2.5, markersize=8, label='신규전액환불률(%)')

    # 데이터 레이블
    for i, v in enumerate(renew_rates):
        if v is not None:
            ax.annotate(f'{v:.2f}%', (i, v), textcoords='offset points', xytext=(0, 12),
                       ha='center', fontsize=11, color=color_renew, fontweight='bold',
                       fontproperties=fm.FontProperties(fname=font_path) if font_path else None)
    for i, v in enumerate(refund_rates):
        if v is not None:
            ax.annotate(f'{v:.2f}%', (i, v), textcoords='offset points', xytext=(0, 12),
                       ha='center', fontsize=11, color=color_refund, fontweight='bold',
                       fontproperties=fm.FontProperties(fname=font_path) if font_path else None)

    ax.set_xticks(x)
    fp_obj = fm.FontProperties(fname=font_path) if font_path else None
    ax.set_xticklabels(week_labels, fontproperties=fp_obj, fontsize=12)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:.2f}%'))
    ax.set_ylim(0, max(max(r for r in renew_rates if r) * 1.3, 30))
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    legend = ax.legend(prop=fp_obj if fp_obj else None, fontsize=11, loc='upper right')

    # X축 레이블
    ax.tick_params(axis='x', labelsize=12)

    # 하단 <도표> 텍스트
    fig.text(0.5, 0.01, '<도표>', ha='center', fontsize=10, color='gray',
             fontproperties=fp_obj)

    plt.tight_layout(rect=[0, 0.03, 1, 1])
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f'  차트 저장: {output_path}')

# ─── 슬라이드별 데이터 채우기 ──────────────────

def pct_str(val):
    if val is None or val == '': return ''
    try:
        f = float(val)
        return f'{f:.2f}%'
    except:
        return str(val)

def num_str(val):
    if val is None or val == '': return ''
    try:
        return str(int(float(val)))
    except:
        return str(val)

# ─── SLIDE 1: 주요지표현황 요약 ──────────────────
def fill_slide1(slide_path, d, a_ns):
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if len(tables) < 3:
        print('  slide1: 테이블 부족'); return

    # 테이블0: 주차 재수강률
    t0 = tables[0]
    rows = t0.findall(f'.//{{{a_ns}}}tr')
    data0 = d.get('slide1_weekly', {})
    # rows: header, 종료예정, 연장완료, 미연장, 환불, 재수강률, 이탈률
    mapping0 = [None, 'calc_due_pt', 'calc_extend_pt', 'calc_nonrenew_pt', 'calc_refund_pt', 'renew_rate', 'churn_rate']
    for i, (row, key) in enumerate(zip(rows, mapping0)):
        if key is None: continue
        cells = row.findall(f'{{{a_ns}}}tc')
        if len(cells) >= 3:
            val = data0.get(key, '')
            if key in ('renew_rate', 'churn_rate'):
                set_cell_text(cells[2], pct_str(val), a_ns)
            else:
                set_cell_text(cells[2], num_str(val), a_ns)

    # 테이블1: 분기 재수강률
    t1 = tables[1]
    rows1 = t1.findall(f'.//{{{a_ns}}}tr')
    data1 = d.get('slide1_quarterly', {})
    for i, (row, key) in enumerate(zip(rows1, mapping0)):
        if key is None: continue
        cells = row.findall(f'{{{a_ns}}}tc')
        if len(cells) >= 3:
            val = data1.get(key, '')
            if key in ('renew_rate', 'churn_rate'):
                set_cell_text(cells[2], pct_str(val), a_ns)
            else:
                set_cell_text(cells[2], num_str(val), a_ns)

    # 테이블2: 신규전액환불률
    t2 = tables[2]
    rows2 = t2.findall(f'.//{{{a_ns}}}tr')
    data2 = d.get('slide1_new_refund', {})
    keys2 = [None, 'open_pt', 'new_full_churn_pt', 'new_full_churn_rate_weekly',
             'new_full_churn_rate_monthly', 'new_full_churn_rate_quarterly']
    for i, (row, key) in enumerate(zip(rows2, keys2)):
        if key is None: continue
        cells = row.findall(f'{{{a_ns}}}tc')
        if len(cells) >= 3:
            val = data2.get(key, '')
            if 'rate' in key:
                set_cell_text(cells[2], pct_str(val), a_ns)
            else:
                set_cell_text(cells[2], num_str(val), a_ns)

    save_slide(tree, slide_path)
    print('  slide1 완료')

# ─── SLIDE 2: 과목별 지표 (수학만 채움, 영어 빈칸) ──
def fill_slide2(slide_path, d, a_ns):
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return
    t = tables[0]
    rows = t.findall(f'.//{{{a_ns}}}tr')
    data = d.get('slide2_math', {})
    # row0: header, row1: 수학, row2: 영어(빈칸)
    if len(rows) >= 2:
        cells = rows[1].findall(f'{{{a_ns}}}tc')
        vals = [
            pct_str(data.get('renew_rate','')),
            pct_str(data.get('extend_rate','')),
            pct_str(data.get('churn_rate','')),
            pct_str(data.get('mid_refund_rate','')),
            pct_str(data.get('new_full_churn_rate','')),
        ]
        for j, val in enumerate(vals):
            if j+1 < len(cells):
                set_cell_text(cells[j+1], val, a_ns)
    # 영어 행은 빈칸으로 (이미 있는 데이터 유지 or 빈칸)
    if len(rows) >= 3:
        cells = rows[2].findall(f'{{{a_ns}}}tc')
        for j in range(1, len(cells)):
            set_cell_text(cells[j], '', a_ns)
    save_slide(tree, slide_path)
    print('  slide2 완료')

# ─── SLIDE 3: 차트 이미지 교체 ──────────────────
def fill_slide3(slide_path, unpacked_dir, d):
    chart_data = d.get('slide3_trend', {})
    labels = chart_data.get('week_labels', ['3월 3주차','3월 4주차','4월 1주차','4월 2주차'])
    renew  = chart_data.get('renew_rates', [None, None, None, None])
    refund = chart_data.get('refund_rates', [None, None, None, None])

    # 이미지 생성
    chart_img = '/home/claude/trend_chart.png'
    create_trend_chart(labels, renew, refund, chart_img)

    # image1.png 교체
    target_img = os.path.join(unpacked_dir, 'ppt', 'media', 'image1.png')
    shutil.copy(chart_img, target_img)
    print('  slide3 차트 이미지 교체 완료')

# ─── SLIDE 4-6: 실별 재수강률 ───────────────────
def fill_slide_dept_renew(slide_path, dept_data_list, period_label, a_ns):
    """
    dept_data_list: [{'dept': '초등실', 'calc_due': X, 'calc_extend': X, 'calc_nonrenew': X, 'calc_refund': X, 'renew_rate': X, 'refund_rate': X}, ...]
    """
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return
    t = tables[0]
    rows = t.findall(f'.//{{{a_ns}}}tr')

    for i, dept_name in enumerate(DEPT_ORDER_SILL):
        row_idx = i + 1  # header 제외
        if row_idx >= len(rows): break
        row = rows[row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')

        # 매핑 찾기
        data = next((x for x in dept_data_list if x.get('dept') == dept_name), None)
        # 실장은 항상 고정
        manager = DEPT_MANAGER.get(dept_name, '')
        if len(cells) > 0: set_cell_text(cells[0], dept_name, a_ns)
        if len(cells) > 1: set_cell_text(cells[1], manager, a_ns)

        if data:
            if len(cells) > 2: set_cell_text(cells[2], num_str(data.get('calc_due_pt','')), a_ns)
            if len(cells) > 3: set_cell_text(cells[3], num_str(data.get('calc_extend_pt','')), a_ns)
            if len(cells) > 4: set_cell_text(cells[4], num_str(data.get('calc_nonrenew_pt','')), a_ns)
            if len(cells) > 5: set_cell_text(cells[5], num_str(data.get('calc_refund_pt','')), a_ns)
            if len(cells) > 6: set_cell_text(cells[6], pct_str(data.get('renew_rate','')), a_ns)
            if len(cells) > 7: set_cell_text(cells[7], pct_str(data.get('refund_rate','')), a_ns)
        else:
            for j in range(2, min(8, len(cells))):
                set_cell_text(cells[j], '', a_ns)

    save_slide(tree, slide_path)

# ─── SLIDE 7-9: 실별 신규전액환불률 ─────────────
def fill_slide_dept_refund(slide_path, dept_data_list, a_ns):
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return
    t = tables[0]
    rows = t.findall(f'.//{{{a_ns}}}tr')

    for i, dept_name in enumerate(DEPT_ORDER_SILL):
        row_idx = i + 1
        if row_idx >= len(rows): break
        row = rows[row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')
        data = next((x for x in dept_data_list if x.get('dept') == dept_name), None)
        manager = DEPT_MANAGER.get(dept_name, '')
        if len(cells) > 0: set_cell_text(cells[0], dept_name, a_ns)
        if len(cells) > 1: set_cell_text(cells[1], manager, a_ns)
        if data:
            if len(cells) > 2: set_cell_text(cells[2], num_str(data.get('open_pt','')), a_ns)
            if len(cells) > 3: set_cell_text(cells[3], num_str(data.get('new_full_churn_pt','')), a_ns)
            if len(cells) > 4: set_cell_text(cells[4], pct_str(data.get('new_full_churn_rate','')), a_ns)
        else:
            for j in range(2, min(5, len(cells))):
                set_cell_text(cells[j], '', a_ns)

    save_slide(tree, slide_path)

# ─── SLIDE 10-12: 팀별 재수강률 (순위) ──────────
def fill_slide_team_renew(slide_path, team_data_list, a_ns):
    """주말팀은 항상 '비고' 행에 배치"""
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return
    t = tables[0]
    rows = t.findall(f'.//{{{a_ns}}}tr')

    # 주말팀 제외하고 재수강률 순위 정렬
    non_weekend = [x for x in team_data_list if x.get('dept') not in ('주말팀',)]
    weekend = [x for x in team_data_list if x.get('dept') == '주말팀']
    non_weekend_sorted = sorted(non_weekend, key=lambda x: float(x.get('renew_rate') or 0), reverse=True)

    # 순위 부여 (동률 처리)
    ranked = []
    prev_rate = None; prev_rank = 0
    for i, item in enumerate(non_weekend_sorted):
        rate = float(item.get('renew_rate') or 0)
        if rate == prev_rate:
            ranked.append((prev_rank, item))
        else:
            prev_rank = i + 1
            ranked.append((prev_rank, item))
            prev_rate = rate

    # 데이터 행: rows[1] ~ rows[8] = 8팀, rows[9] = 비고(주말팀)
    for i, (rank, item) in enumerate(ranked):
        row_idx = i + 1
        if row_idx >= len(rows): break
        row = rows[row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')
        dept = item.get('dept', '')
        manager = DEPT_MANAGER.get(dept, '')
        if len(cells) > 0: set_cell_text(cells[0], str(rank), a_ns)
        if len(cells) > 1: set_cell_text(cells[1], dept, a_ns)
        if len(cells) > 2: set_cell_text(cells[2], manager, a_ns)
        if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('calc_due_pt','')), a_ns)
        if len(cells) > 4: set_cell_text(cells[4], num_str(item.get('calc_extend_pt','')), a_ns)
        if len(cells) > 5: set_cell_text(cells[5], num_str(item.get('calc_nonrenew_pt','')), a_ns)
        if len(cells) > 6: set_cell_text(cells[6], num_str(item.get('calc_refund_pt','')), a_ns)
        if len(cells) > 7: set_cell_text(cells[7], pct_str(item.get('renew_rate','')), a_ns)
        if len(cells) > 8: set_cell_text(cells[8], pct_str(item.get('refund_rate','')), a_ns)

    # 비고 행 (주말팀)
    bigo_row_idx = len(ranked) + 1
    if bigo_row_idx < len(rows) and weekend:
        row = rows[bigo_row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')
        item = weekend[0]
        dept = item.get('dept', '주말팀')
        manager = DEPT_MANAGER.get(dept, '')
        if len(cells) > 0: set_cell_text(cells[0], '비고', a_ns)
        if len(cells) > 1: set_cell_text(cells[1], dept, a_ns)
        if len(cells) > 2: set_cell_text(cells[2], manager, a_ns)
        if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('calc_due_pt','')), a_ns)
        if len(cells) > 4: set_cell_text(cells[4], num_str(item.get('calc_extend_pt','')), a_ns)
        if len(cells) > 5: set_cell_text(cells[5], num_str(item.get('calc_nonrenew_pt','')), a_ns)
        if len(cells) > 6: set_cell_text(cells[6], num_str(item.get('calc_refund_pt','')), a_ns)
        if len(cells) > 7: set_cell_text(cells[7], pct_str(item.get('renew_rate','')), a_ns)
        if len(cells) > 8: set_cell_text(cells[8], pct_str(item.get('refund_rate','')), a_ns)

    save_slide(tree, slide_path)

# ─── SLIDE 13-15: 팀별 신규전액환불률 (순위) ─────
def fill_slide_team_refund(slide_path, team_data_list, a_ns):
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return
    t = tables[0]
    rows = t.findall(f'.//{{{a_ns}}}tr')

    non_weekend = [x for x in team_data_list if x.get('dept') != '주말팀']
    weekend = [x for x in team_data_list if x.get('dept') == '주말팀']
    sorted_list = sorted(non_weekend, key=lambda x: float(x.get('new_full_churn_rate') or 0))

    ranked = []
    prev_rate = None; prev_rank = 0
    for i, item in enumerate(sorted_list):
        rate = float(item.get('new_full_churn_rate') or 0)
        if rate == prev_rate:
            ranked.append((prev_rank, item))
        else:
            prev_rank = i + 1
            ranked.append((prev_rank, item))
            prev_rate = rate

    for i, (rank, item) in enumerate(ranked):
        row_idx = i + 1
        if row_idx >= len(rows): break
        row = rows[row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')
        dept = item.get('dept', '')
        manager = DEPT_MANAGER.get(dept, '')
        if len(cells) > 0: set_cell_text(cells[0], str(rank), a_ns)
        if len(cells) > 1: set_cell_text(cells[1], dept, a_ns)
        if len(cells) > 2: set_cell_text(cells[2], manager, a_ns)
        if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('open_pt','')), a_ns)
        if len(cells) > 4: set_cell_text(cells[4], num_str(item.get('new_full_churn_pt','')), a_ns)
        if len(cells) > 5: set_cell_text(cells[5], pct_str(item.get('new_full_churn_rate','')), a_ns)

    bigo_row_idx = len(ranked) + 1
    if bigo_row_idx < len(rows) and weekend:
        row = rows[bigo_row_idx]
        cells = row.findall(f'{{{a_ns}}}tc')
        item = weekend[0]
        dept = item.get('dept', '주말팀')
        manager = DEPT_MANAGER.get(dept, '')
        if len(cells) > 0: set_cell_text(cells[0], '비고', a_ns)
        if len(cells) > 1: set_cell_text(cells[1], dept, a_ns)
        if len(cells) > 2: set_cell_text(cells[2], manager, a_ns)
        if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('open_pt','')), a_ns)
        if len(cells) > 4: set_cell_text(cells[4], num_str(item.get('new_full_churn_pt','')), a_ns)
        if len(cells) > 5: set_cell_text(cells[5], pct_str(item.get('new_full_churn_rate','')), a_ns)

    save_slide(tree, slide_path)

# ─── SLIDE 16-17: 팀원별 재수강률 ──────────────
def fill_slide_mgr_renew(slide16_path, slide17_path, mgr_data_list, a_ns):
    """도래자>0인 팀원만 재수강률 내림차순, 도래자=0이면 '비고', 2페이지에 걸쳐 표시"""
    valid = [x for x in mgr_data_list if float(x.get('calc_due_pt') or 0) > 0]
    bigo  = [x for x in mgr_data_list if float(x.get('calc_due_pt') or 0) == 0]
    sorted_valid = sorted(valid, key=lambda x: float(x.get('renew_rate') or 0), reverse=True)

    ranked = []
    prev_rate = None; prev_rank = 0
    for i, item in enumerate(sorted_valid):
        rate = float(item.get('renew_rate') or 0)
        if rate == prev_rate:
            ranked.append((prev_rank, item))
        else:
            prev_rank = i + 1
            ranked.append((prev_rank, item))
            prev_rate = rate

    # 슬라이드16: 최대 19명, 슬라이드17: 나머지 + 비고
    PAGE1_MAX = 19
    page1_data = ranked[:PAGE1_MAX]
    page2_data = ranked[PAGE1_MAX:]

    def fill_mgr_table(slide_path, data_chunk, bigo_list=None):
        tree = parse_slide(slide_path)
        tables = get_tables(tree)
        if not tables: return
        # 각 테이블에 데이터 채우기
        t = tables[0]
        rows = t.findall(f'.//{{{a_ns}}}tr')
        for i, (rank, item) in enumerate(data_chunk):
            row_idx = i + 1
            if row_idx >= len(rows): break
            row = rows[row_idx]
            cells = row.findall(f'{{{a_ns}}}tc')
            if len(cells) > 0: set_cell_text(cells[0], str(rank), a_ns)
            if len(cells) > 1: set_cell_text(cells[1], item.get('mgr',''), a_ns)
            if len(cells) > 2: set_cell_text(cells[2], num_str(item.get('calc_due_pt','')), a_ns)
            if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('calc_extend_pt','')), a_ns)
            if len(cells) > 4: set_cell_text(cells[4], num_str(item.get('calc_nonrenew_pt','')), a_ns)
            if len(cells) > 5: set_cell_text(cells[5], num_str(item.get('calc_refund_pt','')), a_ns)
            if len(cells) > 6: set_cell_text(cells[6], pct_str(item.get('renew_rate','')), a_ns)
        save_slide(tree, slide_path)

    fill_mgr_table(slide16_path, page1_data)
    fill_mgr_table(slide17_path, page2_data, bigo)
    print('  slide16/17 완료')

# ─── SLIDE 18: 매출액 빈칸 처리 ─────────────────
def fill_slide18_blank(slide_path, a_ns):
    """매출액 슬라이드는 모든 데이터 셀을 빈칸으로"""
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    for t in tables:
        rows = t.findall(f'.//{{{a_ns}}}tr')
        for r_idx, row in enumerate(rows):
            if r_idx == 0: continue  # header 유지
            cells = row.findall(f'{{{a_ns}}}tc')
            for c_idx, cell in enumerate(cells):
                if c_idx == 0: continue  # 순위 유지
                set_cell_text(cell, '', a_ns)
    save_slide(tree, slide_path)
    print('  slide18 빈칸 처리 완료')

# ─── SLIDE 19: 팀원별 이탈률 ────────────────────
def fill_slide19_mgr_churn(slide_path, mgr_data_list, a_ns):
    """팀원별 이탈률: 기초PT, 이탈PT, 누적이탈률 - 이탈률 오름차순"""
    tree = parse_slide(slide_path)
    tables = get_tables(tree)
    if not tables: return

    valid = [x for x in mgr_data_list if float(x.get('basic_pt') or 0) > 0]
    bigo  = [x for x in mgr_data_list if float(x.get('basic_pt') or 0) == 0]
    sorted_valid = sorted(valid, key=lambda x: float(x.get('churn_rate') or 0))

    ranked = []
    prev_rate = None; prev_rank = 0
    for i, item in enumerate(sorted_valid):
        rate = float(item.get('churn_rate') or 0)
        if rate == prev_rate:
            ranked.append((prev_rank, item))
        else:
            prev_rank = i + 1
            ranked.append((prev_rank, item))
            prev_rate = rate

    # 테이블 3개에 나눠서
    PAGE_SIZE = 17
    pages = [ranked[i:i+PAGE_SIZE] for i in range(0, len(ranked), PAGE_SIZE)]

    all_tables = get_tables(tree)
    for t_idx, t in enumerate(all_tables):
        if t_idx >= len(pages): break
        page_data = pages[t_idx]
        rows = t.findall(f'.//{{{a_ns}}}tr')
        for i, (rank, item) in enumerate(page_data):
            row_idx = i + 1
            if row_idx >= len(rows): break
            row = rows[row_idx]
            cells = row.findall(f'{{{a_ns}}}tc')
            if len(cells) > 0: set_cell_text(cells[0], str(rank), a_ns)
            if len(cells) > 1: set_cell_text(cells[1], item.get('mgr',''), a_ns)
            if len(cells) > 2: set_cell_text(cells[2], num_str(item.get('basic_pt','')), a_ns)
            if len(cells) > 3: set_cell_text(cells[3], num_str(item.get('churn_pt','')), a_ns)
            if len(cells) > 4:
                rate = item.get('churn_rate','')
                set_cell_text(cells[4], pct_str(rate) if rate else '', a_ns)

    save_slide(tree, slide_path)
    print('  slide19 완료')

# ─── 메인 실행 ──────────────────────────────────
def main(data_json_path, output_path=None):
    if output_path is None:
        output_path = OUTPUT_PPTX

    print('데이터 로드...')
    d = load_data(data_json_path)
    a_ns = NS['a']

    print('템플릿 언팩...')
    if os.path.exists(UNPACKED_DIR):
        shutil.rmtree(UNPACKED_DIR)
    os.system(f'python /mnt/skills/public/pptx/scripts/office/unpack.py "{PPTX_TEMPLATE}" "{UNPACKED_DIR}" 2>/dev/null')

    slides_dir = os.path.join(UNPACKED_DIR, 'ppt', 'slides')

    print('슬라이드 채우는 중...')

    # Slide 1
    fill_slide1(f'{slides_dir}/slide1.xml', d, a_ns)

    # Slide 2
    fill_slide2(f'{slides_dir}/slide2.xml', d, a_ns)

    # Slide 3 (차트)
    fill_slide3(f'{slides_dir}/slide3.xml', UNPACKED_DIR, d)

    # Slides 4-6: 실별 재수강률 (주차/월/분기)
    for slide_num, period_key in [(4,'weekly'),(5,'monthly'),(6,'quarterly')]:
        data_list = d.get(f'slide_dept_renew_{period_key}', [])
        fill_slide_dept_renew(f'{slides_dir}/slide{slide_num}.xml', data_list, period_key, a_ns)
        print(f'  slide{slide_num} 완료')

    # Slides 7-9: 실별 신규전액환불률
    for slide_num, period_key in [(7,'weekly'),(8,'monthly'),(9,'quarterly')]:
        data_list = d.get(f'slide_dept_refund_{period_key}', [])
        fill_slide_dept_refund(f'{slides_dir}/slide{slide_num}.xml', data_list, a_ns)
        print(f'  slide{slide_num} 완료')

    # Slides 10-12: 팀별 재수강률 순위
    for slide_num, period_key in [(10,'weekly'),(11,'monthly'),(12,'quarterly')]:
        data_list = d.get(f'slide_team_renew_{period_key}', [])
        fill_slide_team_renew(f'{slides_dir}/slide{slide_num}.xml', data_list, a_ns)
        print(f'  slide{slide_num} 완료')

    # Slides 13-15: 팀별 신규전액환불률 순위
    for slide_num, period_key in [(13,'weekly'),(14,'monthly'),(15,'quarterly')]:
        data_list = d.get(f'slide_team_refund_{period_key}', [])
        fill_slide_team_refund(f'{slides_dir}/slide{slide_num}.xml', data_list, a_ns)
        print(f'  slide{slide_num} 완료')

    # Slides 16-17: 팀원별 재수강률
    mgr_renew = d.get('slide_mgr_renew_quarterly', [])
    fill_slide_mgr_renew(f'{slides_dir}/slide16.xml', f'{slides_dir}/slide17.xml', mgr_renew, a_ns)

    # Slide 18: 매출액 빈칸
    fill_slide18_blank(f'{slides_dir}/slide18.xml', a_ns)

    # Slide 19: 팀원별 이탈률
    mgr_churn = d.get('slide_mgr_churn_monthly', [])
    fill_slide19_mgr_churn(f'{slides_dir}/slide19.xml', mgr_churn, a_ns)

    print('패킹...')
    os.system(f'python /mnt/skills/public/pptx/scripts/office/pack.py "{UNPACKED_DIR}" "{output_path}" --original "{PPTX_TEMPLATE}" 2>/dev/null')

    print(f'완료: {output_path}')
    return output_path

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python generate_ppt.py data.json [output.pptx]')
        sys.exit(1)
    out = main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(f'OUTPUT: {out}')
