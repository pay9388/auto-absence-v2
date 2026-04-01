"""나이스 월별 출결 현황 Excel 파서"""
import re
from datetime import date, timedelta

import openpyxl

# 출결구분 → (유형, 결석종류, 조퇴종류)
ATTENDANCE_MAP = {
    '질병결석':    ('결석', '질병',  ''),
    '출석인정결석': ('결석', '인정',  ''),
    '출석인정지각': ('지각',  '',  '인정지각'),
    '출석인정조퇴': ('조퇴',  '',  '인정조퇴'),
    '출석인정결과': ('결과',  '',  '인정결과'),
}


def _parse_date(val):
    """'2026.03.11.(화)' → date(2026, 3, 11)"""
    if not val:
        return None
    m = re.match(r'(\d{4})\.(\d{2})\.(\d{2})\.', str(val).strip())
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _parse_periods(val):
    """'조회,1교시,3교시,7교시,종례' → ('1', '7') — 숫자 교시만 추출"""
    if not val:
        return '', ''
    nums = []
    for part in str(val).split(','):
        m = re.match(r'^(\d+)교시$', part.strip())
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        return '', ''
    return str(min(nums)), str(max(nums))


def _next_weekday(d):
    """d 이후 첫 번째 평일 반환"""
    next_d = d + timedelta(days=1)
    while next_d.weekday() >= 5:
        next_d += timedelta(days=1)
    return next_d


def _merge_consecutive(students):
    """같은 학생의 연속된 결석(평일 기준)을 한 항목으로 병합"""
    def sort_key(s):
        try:
            return (int(s['번호']), s['시작일'])
        except (ValueError, TypeError):
            return (0, s['시작일'])

    sorted_s = sorted(students, key=sort_key)
    merged = []
    i = 0
    while i < len(sorted_s):
        s = dict(sorted_s[i])
        if s['유형'] == '결석':
            j = i + 1
            while j < len(sorted_s):
                nx = sorted_s[j]
                if (nx['번호'] == s['번호'] and
                        nx['이름'] == s['이름'] and
                        nx['결석종류'] == s['결석종류'] and
                        nx['사유'] == s['사유'] and
                        nx['시작일'] == _next_weekday(s['종료일'])):
                    s['종료일'] = nx['종료일']
                    j += 1
                else:
                    break
            i = j
        else:
            i += 1
        merged.append(s)
    return merged


def parse_neis(filepath):
    """나이스 Excel → 학생 dict 리스트 반환 (builder 호환 형식)"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    raw_header = next(rows_iter)
    header = [str(c).strip() if c is not None else '' for c in raw_header]

    def col(row, name):
        try:
            idx = header.index(name)
            val = row[idx]
            return str(val).strip() if val is not None else ''
        except (ValueError, IndexError):
            return ''

    students = []
    for row in rows_iter:
        if not any(c is not None and str(c).strip() not in ('', 'None') for c in row):
            continue

        출결구분 = col(row, '출결구분')
        if 출결구분 not in ATTENDANCE_MAP:
            continue

        사유 = col(row, '사유')
        if 출결구분 == '출석인정결석' and '현장체험학습' in 사유:
            continue

        유형, 결석종류, 조퇴종류 = ATTENDANCE_MAP[출결구분]
        시작일 = _parse_date(col(row, '일자'))
        if 시작일 is None:
            continue

        시작교시, 종료교시 = '', ''
        if 유형 != '결석':
            시작교시, 종료교시 = _parse_periods(col(row, '결시교시'))

        students.append({
            '번호':    col(row, '번호'),
            '이름':    col(row, '성명'),
            '학부모':  '',
            '유형':    유형,
            '시작일':  시작일,
            '종료일':  시작일,
            '시작교시': 시작교시,
            '종료교시': 종료교시,
            '결석종류': 결석종류,
            '조퇴종류': 조퇴종류,
            '사유':    사유,
            '증빙서류': col(row, '증빙서류'),
        })

    wb.close()
    # 정렬만 하고 병합은 UI에서 수동으로
    def sort_key(s):
        try:
            return (int(s['번호']), s['시작일'])
        except (ValueError, TypeError):
            return (0, s['시작일'])
    return sorted(students, key=sort_key)
