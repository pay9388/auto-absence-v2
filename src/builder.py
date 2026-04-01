"""결석신고서 HWPX 빌더
template.hwpx를 기반으로 여러 학생의 결석신고서를 1파일 다페이지로 생성
"""
import copy
import io
import os
import re
import sys
import zipfile
from datetime import timedelta

from lxml import etree

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'


# ── XML 유틸 ─────────────────────────────────────────────────────────────────

def _tag(name):
    return '{%s}%s' % (HP, name)


def _find_t(el, substr):
    for node in el.iter(_tag('t')):
        if node.text and substr in node.text:
            return node
    return None


def _find_all_t(el, substr):
    return [n for n in el.iter(_tag('t')) if n.text and substr in n.text]


def _check_replace(text, kw):
    """□ keyword → ▣ keyword"""
    return text.replace('□ ' + kw, '▣ ' + kw)


def _set_or_add_text(cell_el, text):
    run = cell_el.find('.//' + _tag('run'))
    if run is None:
        return
    t_el = run.find(_tag('t'))
    if t_el is None:
        t_el = etree.SubElement(run, _tag('t'))
    t_el.text = text


def _count_weekdays(start, end):
    """start ~ end 사이 평일(월~금) 수"""
    count = 0
    current = start
    while current <= end:
        if current.weekday() < 5:
            count += 1
        current += timedelta(days=1)
    return count


# ── 학생 1명분 XML 수정 ───────────────────────────────────────────────────────

def _modify_para(para, student, grade, class_num, teacher_name):
    tbl = para.find('.//' + _tag('tbl'))
    trs = [r for r in tbl if r.tag == _tag('tr')]
    유형 = student['유형']
    시작 = student['시작일']
    종료 = student.get('종료일') or 시작
    보고일 = 종료 + timedelta(days=3)

    # ── tr[2]: 학반 / 성명 ───────────────────────────────
    cells = list(trs[2])
    n = _find_t(cells[1], '학년')
    if n is not None:
        n.text = '%s학년  %s반  %s번' % (grade, class_num, student['번호'])
    _set_or_add_text(cells[3], student['이름'])

    # ── tr[3]: 결석 기간 ─────────────────────────────────
    cells = list(trs[3])
    n = _find_t(cells[2], '년')
    if n is not None:
        if 유형 == '결석':
            days = _count_weekdays(시작, 종료)
            n.text = '%d년  %d월  %d일 ~   %d월   %d일 ( %d )일간' % (
                시작.year, 시작.month, 시작.day, 종료.month, 종료.day, days)
        else:
            n.text = ''

    # ── tr[4]: 지각/조퇴 기간 ────────────────────────────
    cells = list(trs[4])
    n = _find_t(cells[1], '년')
    if n is not None:
        if 유형 != '결석':
            s교시 = student.get('시작교시', '')
            e교시 = student.get('종료교시', '')
            n.text = '%d년  %2d월  %2d일        %s교시 ~      %s교시' % (
                시작.year, 시작.month, 시작.day, s교시, e교시)
        else:
            n.text = ''

    # ── tr[5]: 결석 종류 체크박스 ────────────────────────
    cells = list(trs[5])
    n = _find_t(cells[2], '인정결석')
    if n is not None:
        text = '□ 인정결석 □ 질병결석 □ 기타결석'
        if 유형 == '결석':
            kw_map = {'인정': '인정결석', '질병': '질병결석', '기타': '기타결석'}
            kw = kw_map.get(student.get('결석종류', '').strip(), '')
            if kw:
                text = _check_replace(text, kw)
        n.text = text

    # ── tr[6]: 지각/조퇴 종류 체크박스 ──────────────────
    cells = list(trs[6])
    jt = student.get('조퇴종류', '').strip()
    if 유형 != '결석' and jt:
        for n in cells[1].iter(_tag('t')):
            if n.text and ('□ ' + jt) in n.text:
                n.text = _check_replace(n.text, jt)

    # ── tr[7]: 사유 ──────────────────────────────────────
    cells = list(trs[7])
    _set_or_add_text(cells[1], student.get('사유', ''))

    # ── tr[8]: 증빙서류 체크박스 ─────────────────────────
    cells = list(trs[8])
    jeungbi = student.get('증빙서류', '').strip()
    check_map = {
        '진단서':       '진단서',
        '소견서':       '진단서',
        '진료확인서':   '진단서',
        '진료 확인서':  '진단서',
        '처방전':       '진단서',
        '공문':         '관련 공문',
        '담임교사확인서': '담임교사 확인서',
        '담임교사 확인서': '담임교사 확인서',
        '기타':         '기타:',
    }
    for n in cells[1].iter(_tag('t')):
        if n.text:
            for csv_v, xml_kw in check_map.items():
                if jeungbi == csv_v and ('□ ' + xml_kw) in n.text:
                    n.text = _check_replace(n.text, xml_kw)

    # ── tr[9]: 신고일 / 학생 서명 / 학부모 서명 ──────────
    cells = list(trs[9])
    n = _find_t(cells[0], '년')
    if n is not None:
        n.text = '%d년  %d월  %d일' % (보고일.year, 보고일.month, 보고일.day)
    n = _find_t(cells[0], '학  생:')
    if n is not None:
        n.text = '학  생:  %s   (서명)' % student['이름']

    # ── tr[10]: 담임교사 확인일 + 이름 ───────────────────
    cells = list(trs[10])
    n = _find_t(cells[0], '년')
    if n is not None:
        n.text = '%d년  %d월  %d일' % (보고일.year, 보고일.month, 보고일.day)
    if teacher_name:
        n2 = _find_t(tbl, '(인)')
        if n2 is not None:
            n2.text = re.sub(r'담\s*임:.*', '담  임:  %s  (인)' % teacher_name, n2.text)


# ── 리소스 경로 ───────────────────────────────────────────────────────────────

def _resource(*parts):
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


# ── 메인 빌더 ─────────────────────────────────────────────────────────────────

def build_hwpx(students, grade, class_num, teacher_name, output_path, progress_cb=None):
    """여러 학생 → 1인 1페이지 HWPX 파일 생성

    Args:
        students:     학생 dict 리스트 (parser.parse_neis 반환값)
        grade:        학년 (str)
        class_num:    반 (str)
        teacher_name: 담임교사 이름 (str)
        output_path:  저장 경로 (.hwpx)
        progress_cb:  (current, total) 콜백 (선택)
    """
    tmpl_path = _resource('assets', 'template.hwpx')

    with zipfile.ZipFile(tmpl_path, 'r') as z:
        files = {name: z.read(name) for name in z.namelist()}

    HH = 'http://www.hancom.co.kr/hwpml/2011/head'
    hroot = etree.fromstring(files['Contents/header.xml'])
    for para_pr in hroot.iter('{%s}paraPr' % HH):
        if para_pr.get('id') == '23':
            for ls in para_pr.iter('{%s}lineSpacing' % HH):
                ls.set('type', 'FIXED')
                ls.set('value', '100')
    files['Contents/header.xml'] = etree.tostring(
        hroot, xml_declaration=True, encoding='UTF-8', standalone=True)

    root = etree.fromstring(files['Contents/section0.xml'])
    children = list(root)
    tbl_para = children[1]
    root.remove(tbl_para)

    sec_para = children[0]
    for lineseg in sec_para.iter(_tag('lineseg')):
        lineseg.set('vertsize', '1')
        lineseg.set('textheight', '1')
        lineseg.set('baseline', '1')
        lineseg.set('spacing', '0')

    total = len(students)
    for i, student in enumerate(students):
        if progress_cb:
            progress_cb(i, total)
        clone = copy.deepcopy(tbl_para)
        if i > 0:
            clone.set('pageBreak', '1')
        try:
            _modify_para(clone, student, grade, class_num, teacher_name)
        except Exception as e:
            raise RuntimeError(
                f'[{i+1}번째 학생 "{student.get("이름", "?")}"] 처리 오류: {e}'
            ) from e
        root.append(clone)

    if progress_cb:
        progress_cb(total, total)

    new_xml = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['Contents/section0.xml'] = new_xml

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            if name == 'mimetype':
                zout.writestr(zipfile.ZipInfo(name), data,
                              compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(name, data)
    buf.seek(0)
    with open(output_path, 'wb') as f:
        f.write(buf.read())
