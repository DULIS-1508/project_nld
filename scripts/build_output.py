import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

m = json.load(open('/home/user/project_nld/data/master.json'))

def agg_disability(dc_list):
    if not dc_list:
        return dict(유무=None, 학부계=None, 대학원계=None, 전체재학생=None, 캠퍼스수=0)
    has_center = any(d.get('지원센터유무') == 'O' for d in dc_list)
    any_x = any(d.get('지원센터유무') == 'X' for d in dc_list)
    유무 = 'O' if has_center else ('X' if any_x else None)
    학부계 = sum((d.get('장애재학생수_학부') or 0) for d in dc_list)
    대학원계 = sum((d.get('장애재학생수_대학원') or 0) for d in dc_list)
    전체 = sum((d.get('전체재학생수_학부') or 0) for d in dc_list)
    return dict(유무=유무, 학부계=학부계, 대학원계=대학원계, 전체재학생=전체, 캠퍼스수=len(dc_list))

rows = []
for name, v in m.items():
    agg = agg_disability(v['disability_center'])
    libs = v['libraries']
    비고 = []
    if not libs:
        비고.append('KERIS 도서관 통계 없음')
    if not v['disability_center']:
        비고.append('장애학생지원센터 조사 데이터 없음(전문대학/대학원대학 등 조사 범위 밖)')
    if len(libs) > 1:
        비고.append(f'대학도서관 {len(libs)}개 캠퍼스 통계 존재')
    if agg['캠퍼스수'] > 1:
        비고.append(f'장애학생지원센터 조사 {agg["캠퍼스수"]}개 캠퍼스로 분할 집계')
    if v['비고']:
        비고.append(f'[교육부고시비고] {v["비고"]}')

    rows.append({
        '학교명': name,
        '본분교구분': v['본분교구분'],
        '대학구분': v['대학구분'],
        '학교구분': v['학교구분'],
        '지역': v['지역'],
        '설립구분': v['설립구분'],
        '대학도서관_캠퍼스수': len(libs),
        '대학도서관명(대표)': libs[0]['lib_row_name'] if libs else '',
        '대학도서관명(전체)': '; '.join(l['lib_row_name'] for l in libs),
        '장애학생지원센터_유무': agg['유무'] if agg['유무'] else ('조사없음' if not v['disability_center'] else ''),
        '전체재학생수(학부,합계)': agg['전체재학생'],
        '장애재학생수(학부,합계)': agg['학부계'],
        '장애재학생수(대학원,합계)': agg['대학원계'],
        '설문대상포함': 'O',  # per user decision: all MOE-gazetted schools included
        '비고': ' / '.join(비고),
    })

# sort: 지원센터 있음 먼저, 그다음 학교구분, 그다음 이름
gubun_order = {'대학교':0,'산업대학':1,'교육대학':2,'방송통신대학':3,'기술대학':4,'각종학교(대학)':5,
               '사이버대학':6,'사이버대학(전문)':7,'전문대학':8,'전문대학원':9,'특수대학원':10,
               '기능대학':11,'일반대학원':12}
rows.sort(key=lambda r: (gubun_order.get(r['학교구분'], 99), r['학교명']))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '대상기관 목록'

headers = list(rows[0].keys())
ws.append(headers)
for r in rows:
    ws.append([r[h] for h in headers])

# styling
header_fill = PatternFill(start_color='2B4A73', end_color='2B4A73', fill_type='solid')
header_font = Font(color='FFFFFF', bold=True)
thin = Side(style='thin', color='CCCCCC')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws.freeze_panes = 'A2'
widths = [22,10,10,14,8,10,14,26,40,16,14,14,16,12,50]
for i,w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(vertical='center', wrap_text=(cell.column_letter in ('I','O')))

# ---- Sheet 2: 캠퍼스별 상세 (multi-campus schools) ----
ws2 = wb.create_sheet('캠퍼스별 상세')
ws2.append(['학교명','구분','캠퍼스/도서관명','출처','지역/캠퍼스태그','비고'])
for cell in ws2[1]:
    cell.fill = header_fill; cell.font = header_font
for name, v in m.items():
    if len(v['libraries']) > 1:
        for l in v['libraries']:
            ws2.append([name, '대학도서관', l['lib_row_name'], 'KERIS', l['본분교'], l['match_method']])
    if len(v['disability_center']) > 1:
        for d in v['disability_center']:
            ws2.append([name, '장애학생지원센터', d['raw_name'], '장애학생지원센터 현황조사', d['campus_tag'],
                        f"장애재학생(학부) {d['장애재학생수_학부']}명, 지원센터유무 {d['지원센터유무']}"])
ws2.column_dimensions['A'].width = 22
ws2.column_dimensions['C'].width = 30
ws2.column_dimensions['F'].width = 40

# ---- Sheet 3: 제외/특이사항 로그 ----
ws3 = wb.create_sheet('제외 및 특이사항')
ws3.append(['구분','원본 기재명','처리 결과','사유'])
for cell in ws3[1]:
    cell.fill = header_fill; cell.font = header_font
log_rows = [
    ('KERIS 도서관 통계', '국방대학교 도서관', '미포함(교육부 고시 목록에 없음)', '국방대학교는 「국방대학교 설치법」 소관 특수대학원으로 대학알리미 공시대상학교 목록에 미등재. 설문 대상 포함 여부 별도 확인 필요'),
    ('KERIS 도서관 통계', '육군사관학교 학술정보원', '미포함(교육부 고시 목록에 없음)', '사관학교는 별도 소관법령(육군3사관학교설치법 등) 대상으로 대학알리미 공시대상학교 목록에 미등재. 설문 대상 포함 여부 별도 확인 필요'),
    ('KERIS 도서관 통계', '국제예술대학교 도서관', '제외', '평생교육법상 학교형태 평생교육시설(전공대학)로, 고등교육법 제2조 학교가 아니어서 공시대상 아님'),
    ('장애학생지원센터 현황조사', '경남과학기술대학교_본교', '제외(중복 데이터)', '2021년 경상대학교와 통합되어 경상국립대학교로 개편됨. 조사원본에 옛 명칭 데이터가 남아있으나 경상국립대학교 행에 현재 데이터가 별도로 존재'),
    ('교육부 고시 목록 반영(신뢰도 검증)', '경남도립거창대학 / 경남도립남해대학', '국립창원대학교로 매칭', '2026.3. 국립창원대학교로 통합(글로컬대학 사업)'),
    ('교육부 고시 목록 반영(신뢰도 검증)', '서라벌대학교', '신경주대학교로 매칭', '2024.3. 경주대학교와 통합, 신경주대학교로 개편'),
    ('교육부 고시 목록 반영(신뢰도 검증)', '원광보건대학교', '원광대학교로 매칭', '2026.2. 원광대학교로 흡수통합'),
    ('교육부 고시 목록 반영(신뢰도 검증)', '국립강릉원주대학교', '강원대학교로 매칭', '2026.3. 강원대학교와 통합(통합 강원대)'),
    ('교육부 고시 목록 반영(신뢰도 검증)', '전남도립대학교', '국립목포대학교로 매칭', '2026.3. 국립목포대학교와 통합(2·4년제 통합모델)'),
    ('명칭 표기 차이(자동 정정)', '한국폴리텍 각 캠퍼스', '한국폴리텍 I~VII대학/특성화대학으로 매칭', 'KERIS는 로마숫자 유니코드 문자(Ⅰ,Ⅱ...) 및 붙여쓰기로 표기, 교육부 고시는 로마자(I, II...)+띄어쓰기로 표기하여 자동매칭 실패, 수동 보정'),
]
for r in log_rows:
    ws3.append(r)
ws3.column_dimensions['A'].width = 26
ws3.column_dimensions['B'].width = 26
ws3.column_dimensions['C'].width = 30
ws3.column_dimensions['D'].width = 70
for row in ws3.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(vertical='top', wrap_text=True)

out_path = '/home/user/project_nld/data/대학도서관_장애학생지원센터_대상기관목록.xlsx'
wb.save(out_path)
print('saved', out_path)
print('total institutions:', len(rows))
from collections import Counter
print(Counter(r['장애학생지원센터_유무'] for r in rows))
print(Counter(r['학교구분'] for r in rows))
