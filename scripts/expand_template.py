"""
expand_template.py
application.hwpx 템플릿의 2·3차년도 Gantt 행을 미리 9개씩 확장.
이후 write_hwpx.py에서 행 클로닝 없이 텍스트만 채우면 됨.
"""
import zipfile, os, copy, shutil
from lxml import etree

HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'

SKILL_DIR = os.path.expanduser('~/.claude/skills/k-grant-apply')
TEMPLATE   = os.path.join(SKILL_DIR, 'assets/application.hwpx')
BACKUP     = TEMPLATE + '.bak'
TARGET_YEAR2_ROWS = 9
TARGET_YEAR3_ROWS = 9


# ── helpers ────────────────────────────────────────────────────────────

def get_text(elem):
    return ''.join(t.text or '' for t in elem.iter(f'{HP}t')).strip()


def make_empty_row(template_row):
    """deepcopy template row, clear all text & linesegarray."""
    new_row = copy.deepcopy(template_row)
    for t in new_row.iter(f'{HP}t'):
        t.text = ''
    for lsa in new_row.findall(f'.//{HP}linesegarray'):
        par = lsa.getparent()
        if par is not None:
            par.remove(lsa)
    return new_row


def renumber_row_addrs(tbl):
    """Set cellAddr.rowAddr = actual row index for every cell in tbl."""
    direct_rows = [c for c in tbl if c.tag == f'{HP}tr']
    for row_idx, row in enumerate(direct_rows):
        for tc in row:
            if tc.tag != f'{HP}tc':
                continue
            addr = tc.find(f'{HP}cellAddr')
            if addr is not None:
                addr.set('rowAddr', str(row_idx))


def recalc_heights(inner_tbl, outer_tbl):
    """Recalculate sz.height for inner and outer tables."""
    direct_rows = [c for c in inner_tbl if c.tag == f'{HP}tr']
    total_h = 0
    for row in direct_rows:
        cells = [c for c in row if c.tag == f'{HP}tc']
        if cells:
            csz = cells[0].find(f'{HP}cellSz')
            if csz is not None:
                total_h += int(csz.get('height', 0))

    sz_i = inner_tbl.find(f'{HP}sz')
    sz_o = outer_tbl.find(f'{HP}sz') if outer_tbl is not None else None

    if sz_i is not None:
        old_inner = int(sz_i.get('height', 0))
        sz_i.set('height', str(total_h))
        if sz_o is not None:
            old_outer = int(sz_o.get('height', 0))
            delta = total_h - old_inner
            sz_o.set('height', str(old_outer + delta))

    print(f'  inner sz.height: {sz_i.get("height") if sz_i is not None else "?"}')
    print(f'  outer sz.height: {sz_o.get("height") if sz_o is not None else "?"}')


# ── main ────────────────────────────────────────────────────────────────

def expand_template():
    # 1. Backup
    shutil.copy2(TEMPLATE, BACKUP)
    print(f'Backup: {BACKUP}')

    # 2. Read template
    with zipfile.ZipFile(TEMPLATE, 'r') as zin:
        entries = {info.filename: (info, zin.read(info.filename))
                   for info in zin.infolist()}

    section_xml = entries['Contents/section0.xml'][1]
    root = etree.fromstring(section_xml)

    # 3. Find outer + inner Gantt table
    for p in root.iter(f'{HP}p'):
        if '4) 연구개발과제 수행일정' in get_text(p):
            sched_p = p
            break

    parent = sched_p.getparent()
    siblings = list(parent)
    for sib in siblings[siblings.index(sched_p) + 1:]:
        tbls = list(sib.iter(f'{HP}tbl'))
        if tbls:
            outer_tbl = tbls[0]
            inner_tbl = tbls[-1]
            break

    direct_rows = [c for c in inner_tbl if c.tag == f'{HP}tr']
    print(f'Template rows before: {len(direct_rows)}')

    # 4. Identify year sections
    # Structure: ..., row[30]="2차 년도", row[31..33]=year2 tasks,
    #            row[34]="N차년도",        row[35..36]=year3 tasks
    year2_header_idx = None
    year3_header_idx = None
    year2_task_rows  = []   # row indices
    year3_task_rows  = []

    for i, row in enumerate(direct_rows):
        cells = [c for c in row if c.tag == f'{HP}tc']
        if not cells:
            continue
        txt = get_text(cells[0]).strip()
        if len(cells) == 1 and '2차 년도' in txt:
            year2_header_idx = i
        elif len(cells) == 1 and ('N차년도' in txt or '3차년도' in txt):
            year3_header_idx = i
        elif year2_header_idx is not None and year3_header_idx is None:
            year2_task_rows.append(i)
        elif year3_header_idx is not None:
            year3_task_rows.append(i)

    print(f'  year2 header @ row[{year2_header_idx}], task rows: {year2_task_rows}')
    print(f'  year3 header @ row[{year3_header_idx}], task rows: {year3_task_rows}')

    # 5. Expand year2 to TARGET_YEAR2_ROWS
    # Insert new rows after last year2 task row (before year3 header)
    year2_template_row = direct_rows[year2_task_rows[-1]]
    year2_add = TARGET_YEAR2_ROWS - len(year2_task_rows)
    if year2_add > 0:
        last_row = year2_template_row
        for _ in range(year2_add):
            new_row = make_empty_row(year2_template_row)
            idx = list(inner_tbl).index(last_row)
            inner_tbl.insert(idx + 1, new_row)
            last_row = new_row
        print(f'  Added {year2_add} rows for year2')

    # Refresh after insertions
    direct_rows = [c for c in inner_tbl if c.tag == f'{HP}tr']

    # Re-find year3 header (index shifted)
    year3_header_idx_new = None
    year3_task_rows_new = []
    for i, row in enumerate(direct_rows):
        cells = [c for c in row if c.tag == f'{HP}tc']
        if not cells:
            continue
        txt = get_text(cells[0]).strip()
        if len(cells) == 1 and ('N차년도' in txt or '3차년도' in txt):
            year3_header_idx_new = i
        elif year3_header_idx_new is not None:
            year3_task_rows_new.append(i)

    print(f'  year3 header now @ row[{year3_header_idx_new}], task rows: {year3_task_rows_new}')

    # 6. Expand year3 to TARGET_YEAR3_ROWS
    year3_template_row = direct_rows[year3_task_rows_new[-1]]
    year3_add = TARGET_YEAR3_ROWS - len(year3_task_rows_new)
    if year3_add > 0:
        last_row = year3_template_row
        for _ in range(year3_add):
            new_row = make_empty_row(year3_template_row)
            idx = list(inner_tbl).index(last_row)
            inner_tbl.insert(idx + 1, new_row)
            last_row = new_row
        print(f'  Added {year3_add} rows for year3')

    # Refresh
    direct_rows = [c for c in inner_tbl if c.tag == f'{HP}tr']
    print(f'Template rows after expansion: {len(direct_rows)}')

    # 7. Update rowCnt attribute
    inner_tbl.set('rowCnt', str(len(direct_rows)))
    print(f'rowCnt updated: {len(direct_rows)}')

    # 8. Renumber cellAddr.rowAddr
    renumber_row_addrs(inner_tbl)
    print('cellAddr renumbered')

    # 8. Recalculate table heights
    recalc_heights(inner_tbl, outer_tbl)

    # 9. Save new template
    new_section = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(TEMPLATE, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fn, (info, content) in entries.items():
            if fn == 'Contents/section0.xml':
                zout.writestr(info, new_section)
            elif fn == 'mimetype':
                zout.writestr(info, content, compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(info, content)

    print(f'\n템플릿 저장 완료: {TEMPLATE}')


if __name__ == '__main__':
    expand_template()
