#!/usr/bin/env python3
"""
write_hwpx.py - NRF 연구개발과제 연구계획서 HWPX 생성 스크립트

HWPX 템플릿의 "○     -" 플레이스홀더를 실제 연구 내용으로 교체한다.
각 섹션의 내용은 JSON 파일로 전달한다.

Usage:
    python3 write_hwpx.py --template <template.hwpx> --output <output.hwpx> --data-json <data.json>
"""
import argparse
import copy
import json
import os
import re
import sys
import zipfile

from lxml import etree

HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
HS = '{http://www.hancom.co.kr/hwpml/2011/section}'


# =========================================================================
# Helper Functions
# =========================================================================

def get_all_text(elem):
    """Get all text from <hp:t> elements."""
    texts = []
    for t in elem.iter(HP + 't'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def strip_linesegarray(p):
    """Remove layout cache to force re-wrap."""
    for lsa in p.findall(HP + 'linesegarray'):
        p.remove(lsa)


def make_content_para(template_p, text):
    """Deep-copy a template paragraph and set its text.

    Strips embedded tables (작성요령 boxes) from the copy,
    keeps only text runs.
    """
    new_p = copy.deepcopy(template_p)
    # Remove embedded table runs (작성요령 box)
    for run in list(new_p.findall(HP + 'run')):
        if run.find(HP + 'tbl') is not None:
            new_p.remove(run)
        elif run.find('.//' + HP + 'tbl') is not None:
            new_p.remove(run)
    strip_linesegarray(new_p)

    # Set text in the first run that has <hp:t>
    for run in new_p.findall(HP + 'run'):
        t_elem = run.find(HP + 't')
        if t_elem is not None:
            for child in list(t_elem):
                t_elem.remove(child)
            t_elem.text = text
            return new_p
        else:
            t_elem = etree.SubElement(run, HP + 't')
            t_elem.text = text
            return new_p

    # If no run with <hp:t>, add one
    run = etree.SubElement(new_p, HP + 'run')
    run.set('charPrIDRef', '35')
    t_elem = etree.SubElement(run, HP + 't')
    t_elem.text = text
    return new_p


def find_all_placeholder_paras(root):
    """Find all '○     -' placeholder paragraphs in document order.

    Returns list of (paragraph_element, parent_element, index_in_parent).
    """
    placeholders = []
    for p in root.iter(HP + 'p'):
        text = get_all_text(p)
        if text == '○     -':
            parent = p.getparent()
            if parent is not None:
                idx = list(parent).index(p)
                placeholders.append((p, parent, idx))
    return placeholders


def replace_placeholder(placeholder_p, parent, idx, content_lines):
    """Replace a '○     -' placeholder with content lines.

    Args:
        placeholder_p: The placeholder <hp:p> element
        parent: Its parent element
        idx: Its index in parent
        content_lines: List of strings, each becoming a paragraph
                       Format: "○ 내용" or "   - 세부내용"
    """
    if not content_lines:
        # Leave as empty placeholder
        return

    # Insert new paragraphs before the placeholder, then remove it
    insert_idx = idx
    for i, line in enumerate(content_lines):
        new_p = make_content_para(placeholder_p, line)
        parent.insert(insert_idx + i, new_p)

    # Remove the original placeholder
    parent.remove(placeholder_p)


def find_para_by_text(root, search_text):
    """Find the first <hp:p> containing search_text."""
    for p in root.iter(HP + 'p'):
        if search_text in get_all_text(p):
            return p
    return None


def set_para_text(p, new_text):
    """Set text of paragraph (removing embedded tables)."""
    # Remove embedded table runs
    for run in list(p.findall(HP + 'run')):
        if run.find(HP + 'tbl') is not None:
            p.remove(run)
        elif run.find('.//' + HP + 'tbl') is not None:
            p.remove(run)

    strip_linesegarray(p)

    for run in p.findall(HP + 'run'):
        t_elem = run.find(HP + 't')
        if t_elem is not None:
            for child in list(t_elem):
                t_elem.remove(child)
            t_elem.text = new_text
            return True

    # Add run if none exists
    run = etree.SubElement(p, HP + 'run')
    run.set('charPrIDRef', '35')
    t_elem = etree.SubElement(run, HP + 't')
    t_elem.text = new_text
    return True


def set_cell_text(tc, new_text):
    """Set text of a table cell's first paragraph."""
    for p in tc.iter(HP + 'p'):
        set_para_text(p, new_text)
        return True
    return False


# =========================================================================
# Year block expansion (N차년도 → 실제 연차)
# =========================================================================

def expand_year_blocks(root, total_years, stage=1):
    """Expand '□ N차년도(N단계)' blocks into actual year/stage blocks.

    - Replaces '□ N차년도(N단계) 목표:' with '□ 2차년도({stage}단계) 목표:'
    - Replaces '□ N차년도(N단계)' with '□ 2차년도({stage}단계)'
    - Clones block(s) for year 3, 4, ... if total_years >= 3
    """
    _expand_goal_year_blocks(root, total_years, stage)
    _expand_content_year_blocks(root, total_years, stage)
    _fix_all_n_year_text(root, total_years, stage)


def _expand_goal_year_blocks(root, total_years, stage):
    """Handle □ N차년도(N단계) 목표: blocks."""
    target_text = '□ N차년도(N단계) 목표:'
    for p in list(root.iter(HP + 'p')):
        text = get_all_text(p)
        if text != target_text:
            continue

        new_year_text = f'□ 2차년도({stage}단계) 목표:'
        set_para_text(p, new_year_text)

        if total_years < 3:
            continue

        parent = p.getparent()
        if parent is None:
            continue
        idx = list(parent).index(p)
        siblings = list(parent)

        # Collect this year block: □ + ○주관 + ○공동 + ○위탁 (+ optional empty)
        block = [p]
        for j in range(idx + 1, len(siblings)):
            sib = siblings[j]
            st = get_all_text(sib)
            if '□' in st and '차년도' in st:
                break
            block.append(sib)
            if st == '':  # empty separator = end of block
                break

        insert_pos = idx + len(block)
        for yr in range(3, total_years + 1):
            for k, bp in enumerate(block):
                new_p = copy.deepcopy(bp)
                bp_text = get_all_text(bp)
                if '2차년도' in bp_text:
                    set_para_text(new_p, bp_text.replace('2차년도', f'{yr}차년도'))
                parent.insert(insert_pos + k, new_p)
            insert_pos += len(block)
        break  # only one N차년도 target per section


def _expand_content_year_blocks(root, total_years, stage):
    """Handle □ N차년도(N단계) blocks in research content section."""
    target_text = '□ N차년도(N단계)'
    for p in list(root.iter(HP + 'p')):
        text = get_all_text(p)
        if text != target_text:
            continue

        new_year_text = f'□ 2차년도({stage}단계)'
        set_para_text(p, new_year_text)

        if total_years < 3:
            continue

        parent = p.getparent()
        if parent is None:
            continue
        idx = list(parent).index(p)
        siblings = list(parent)

        # Collect this year block until next □ or end
        block = [p]
        for j in range(idx + 1, len(siblings)):
            sib = siblings[j]
            st = get_all_text(sib)
            if '□' in st and '차년도' in st:
                break
            block.append(sib)

        insert_pos = idx + len(block)
        for yr in range(3, total_years + 1):
            for k, bp in enumerate(block):
                new_p = copy.deepcopy(bp)
                bp_text = get_all_text(bp)
                if '2차년도' in bp_text:
                    set_para_text(new_p, bp_text.replace('2차년도', f'{yr}차년도'))
                parent.insert(insert_pos + k, new_p)
            insert_pos += len(block)
        break


def _fix_all_n_year_text(root, total_years, stage):
    """Replace any remaining 'N차년도'/'N단계' in all text nodes."""
    for t in root.iter(HP + 't'):
        if t.text:
            if 'N차년도' in t.text:
                t.text = t.text.replace('N차년도', f'{total_years}차년도')
            if 'N단계' in t.text:
                t.text = t.text.replace('N단계', f'{stage}단계')


# =========================================================================
# Section handlers
# =========================================================================

def fill_yearly_goals(root, data):
    """Fill yearly goal placeholders: '○ (주관연구개발과제):' etc.

    data: dict with keys like 'year1_main', 'year1_joint', 'year1_contracted',
                              'year2_main', 'year2_joint', 'year2_contracted', ...
    """
    goal_section = find_para_by_text(root, '가. 연구개발 목표')
    content_section = find_para_by_text(root, '3) 연구개발과제의 내용')

    year_idx = 0
    in_goal = False
    seen_square = False

    for p in root.iter(HP + 'p'):
        text = get_all_text(p)
        if goal_section is not None and p is goal_section:
            in_goal = True
            continue
        if content_section is not None and p is content_section:
            break
        if not in_goal:
            continue

        if '□' in text and '차년도' in text:
            year_idx += 1
            seen_square = True
            continue
        if seen_square and text == '○ (주관연구개발과제):':
            year_key = f'year{year_idx}_main'
            if year_key in data and data[year_key]:
                set_para_text(p, '○ (주관연구개발과제): ' + data[year_key])
        elif seen_square and text == '○ (공동연구개발과제):':
            year_key = f'year{year_idx}_joint'
            if year_key in data and data[year_key]:
                set_para_text(p, '○ (공동연구개발과제): ' + data[year_key])
        elif seen_square and text == '○ (위탁연구개발과제):':
            year_key = f'year{year_idx}_contracted'
            if year_key in data and data[year_key]:
                set_para_text(p, '○ (위탁연구개발과제): ' + data[year_key])


def _fill_org_content_lines(parent, org_para_idx, items):
    """Fill '- 연구개발 내용' lines after an ○ paragraph, cloning if needed."""
    if not items:
        return

    children = list(parent)
    # Find existing '- 연구개발 내용' lines
    content_line_paras = []
    for j in range(org_para_idx + 1, min(org_para_idx + 20, len(children))):
        st = get_all_text(children[j])
        if st.startswith('- 연구개발 내용') or (content_line_paras and st == ''):
            if st.startswith('- 연구개발 내용'):
                content_line_paras.append(children[j])
            else:
                break
        elif '○' in st or '□' in st:
            break

    if not content_line_paras:
        return

    template_para = content_line_paras[-1]

    # Clone template lines if we need more
    while len(content_line_paras) < len(items):
        new_p = copy.deepcopy(template_para)
        # Find insert position: after last content line para
        last_para = content_line_paras[-1]
        last_idx = list(parent).index(last_para)
        parent.insert(last_idx + 1, new_p)
        children = list(parent)
        content_line_paras.append(children[last_idx + 1])

    # Fill content
    for k, content_para in enumerate(content_line_paras):
        if k < len(items):
            set_para_text(content_para, f'   - {items[k]}')
        else:
            set_para_text(content_para, '')


def fill_yearly_contents(root, data):
    """Fill yearly content placeholders.

    data: dict with keys like 'year1_main', 'year1_joint', 'year1_contracted',
          each value is a list of strings ['내용1', '내용2', ...]
    """
    content_section = find_para_by_text(root, '가. 연구개발 내용')
    schedule_section = find_para_by_text(root, '4) 연구개발과제 수행일정 및 주요 결과물')

    if content_section is None:
        return

    year_idx = 0
    collecting = False
    content_paras = []

    for p in root.iter(HP + 'p'):
        if p is content_section:
            collecting = True
            continue
        if schedule_section is not None and p is schedule_section:
            break
        if collecting:
            content_paras.append(p)

    for p in content_paras:
        text = get_all_text(p)
        if '□' in text and '차년도' in text:
            year_idx += 1
            continue

        if text == '○ (주관연구개발과제):':
            year_key = f'year{year_idx}_main'
            if year_key in data and data[year_key]:
                parent = p.getparent()
                if parent is not None:
                    idx = list(parent).index(p)
                    _fill_org_content_lines(parent, idx, data[year_key])

        elif text == '○ (공동연구개발과제):':
            year_key = f'year{year_idx}_joint'
            if year_key in data and data[year_key]:
                parent = p.getparent()
                if parent is not None:
                    idx = list(parent).index(p)
                    _fill_org_content_lines(parent, idx, data[year_key])

        elif text == '○ (위탁연구개발과제):':
            year_key = f'year{year_idx}_contracted'
            if year_key in data and data[year_key]:
                parent = p.getparent()
                if parent is not None:
                    idx = list(parent).index(p)
                    _fill_org_content_lines(parent, idx, data[year_key])


# =========================================================================
# Schedule table filling
# =========================================================================

def fill_schedule_table(root, schedule_data, total_years):
    """Fill schedule table task names and results.

    schedule_data: dict with 'year1', 'year2', 'year3' keys,
                   each a list of {'task': str, 'result': str}
    total_years: int (to fix N차년도 labels)
    """
    schedule_section = find_para_by_text(root, '4) 연구개발과제 수행일정 및 주요 결과물')
    if schedule_section is None:
        return

    # Find the table paragraph after schedule_section
    parent = schedule_section.getparent()
    if parent is None:
        return

    siblings = list(parent)
    sec_idx = siblings.index(schedule_section)

    for sib in siblings[sec_idx + 1:]:
        # Look for embedded table in this paragraph
        for tbl in sib.iter(HP + 'tbl'):
            _fill_schedule_tbl(tbl, schedule_data, total_years)
            return


def _fill_schedule_tbl(tbl, schedule_data, total_years):
    """Fill rows in the schedule Gantt table."""
    year_label_map = {
        '1차 년도': 'year1',
        '2차 년도': 'year2',
        f'{total_years}차년도': f'year{total_years}',
        'N차년도': f'year{total_years}',
    }
    # Also handle generic year3
    year_label_map['3차년도'] = 'year3'
    year_label_map['3차 년도'] = 'year3'

    rows = list(tbl.findall('.//' + HP + 'tr'))
    current_year_key = None
    task_row_idx = 0
    is_header_row = False

    for row in rows:
        cells = list(row.findall('.//' + HP + 'tc'))
        if not cells:
            continue

        cell0_text = get_all_text(cells[0]).strip()

        # Detect year label rows
        matched_year = None
        for label, key in year_label_map.items():
            if label in cell0_text or cell0_text == label:
                matched_year = key
                break

        if matched_year:
            current_year_key = matched_year
            task_row_idx = 0
            is_header_row = True
            # Fix N차년도 in cell text
            if 'N차년도' in cell0_text:
                set_cell_text(cells[0], cell0_text.replace('N차년도', f'{total_years}차년도'))
            continue

        # Skip header rows (추진내용, 1, 2, 3, ... header)
        if cell0_text in ('추진내용', '추진 일정', '결과물', '1', '2', '3',
                          '4', '5', '6', '7', '8', '9', '10', '11', '12'):
            is_header_row = True
            continue

        # Empty task rows
        if current_year_key is None:
            continue

        tasks = schedule_data.get(current_year_key, [])
        if task_row_idx < len(tasks):
            task = tasks[task_row_idx]
            # First cell = task name
            if cells:
                set_cell_text(cells[0], task.get('task', ''))
            # Last non-empty-header cell = result
            # Find the result column (typically last column)
            if len(cells) > 1:
                set_cell_text(cells[-1], task.get('result', ''))
        task_row_idx += 1


# =========================================================================
# Main modification function
# =========================================================================

def modify_application(section_xml, data):
    """Modify the NRF application HWPX section0.xml.

    data structure:
    {
      "_meta": {
        "total_years": 3,
        "stage": 1
      },
      "necessity": ["단락1", "단락2", ...],
      "final_goal": ["목표1", "목표2", ...],
      "yearly_goals": {
        "year1_main": "주관과제 1차년도 목표",
        "year1_joint": "공동과제 1차년도 목표 (없으면 빈 문자열)",
        "year1_contracted": "",
        "year2_main": "주관과제 2차년도 목표",
        ...
      },
      "yearly_contents": {
        "year1_main": ["연구내용1", "연구내용2", ...],
        ...
      },
      "schedule": {
        "year1": [{"task": "추진내용", "result": "결과물"}, ...],
        "year2": [...],
        "year3": [...]
      },
      "strategy": ["전략1", ...],
      "system": ["체계1", ...],
      "utilization": ["활용1", ...],
      "effects": ["효과1", ...],
      "commercialization": {
        "market_size": [...], ...
      }
    }
    """
    root = etree.fromstring(section_xml)

    # Extract meta info
    meta = data.get('_meta', {})
    total_years = meta.get('total_years', 2)
    stage = meta.get('stage', 1)

    # Step 1: Expand year blocks (N차년도 → actual year numbers, clone for year 3+)
    expand_year_blocks(root, total_years, stage)
    print(f"  Year blocks expanded: {total_years}차년도, {stage}단계", file=sys.stderr)

    # Step 2: Find all placeholder paragraphs
    placeholders = find_all_placeholder_paras(root)
    print(f"Found {len(placeholders)} placeholder paragraphs", file=sys.stderr)

    # Section map: placeholder index → content
    section_map = {
        0: ('necessity', data.get('necessity', [])),
        1: ('final_goal', data.get('final_goal', [])),
        2: ('strategy', data.get('strategy', [])),
        3: ('system', data.get('system', [])),
        4: ('utilization', data.get('utilization', [])),
        5: ('effects', data.get('effects', [])),
    }

    commercialization = data.get('commercialization', {})
    comm_sections = [
        'market_size', 'demand', 'competition', 'ip',
        'standardization', 'biz_strategy', 'investment', 'production'
    ]
    for i, key in enumerate(comm_sections):
        section_map[6 + i] = (key, commercialization.get(key, []))

    placeholder_contents = []
    for i in range(len(placeholders)):
        if i in section_map:
            _, content = section_map[i]
            if isinstance(content, str):
                content = [content] if content else []
            placeholder_contents.append(content)
        else:
            placeholder_contents.append([])

    # Fill placeholders in reverse order
    for i in range(len(placeholders) - 1, -1, -1):
        p, parent, idx = placeholders[i]
        content = placeholder_contents[i]
        if content:
            replace_placeholder(p, parent, idx, content)

    # Step 3: Fill yearly goals
    yearly_goals = data.get('yearly_goals', {})
    if yearly_goals:
        fill_yearly_goals(root, yearly_goals)

    # Step 4: Fill yearly contents (with line cloning)
    yearly_contents = data.get('yearly_contents', {})
    if yearly_contents:
        fill_yearly_contents(root, yearly_contents)

    # Step 5: Fill schedule table
    schedule = data.get('schedule', {})
    if schedule:
        fill_schedule_table(root, schedule, total_years)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# =========================================================================
# HWPX File Writer
# =========================================================================

def write_hwpx(template_path, output_path, data):
    """Copy HWPX template and modify section0.xml."""
    entries = {}
    with zipfile.ZipFile(template_path, 'r') as zin:
        for info in zin.infolist():
            entries[info.filename] = (info, zin.read(info.filename))

    section_xml = entries['Contents/section0.xml'][1]
    new_section = modify_application(section_xml, data)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for filename, (info, content) in entries.items():
            if filename == 'Contents/section0.xml':
                zout.writestr(info, new_section)
            elif filename == 'mimetype':
                zout.writestr(info, content, compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(info, content)

    return output_path


# =========================================================================
# CLI
# =========================================================================

def main():
    parser = argparse.ArgumentParser(description='NRF 연구계획서 HWPX 생성')
    parser.add_argument('--template', required=True, help='HWPX 템플릿 파일 경로')
    parser.add_argument('--output', required=True, help='출력 HWPX 파일 경로')
    parser.add_argument('--data-json', required=True, help='데이터 JSON 파일 경로')
    args = parser.parse_args()

    with open(args.data_json, 'r', encoding='utf-8') as f:
        data = json.load(f)

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)

    result = write_hwpx(args.template, args.output, data)
    print(f'연구계획서 생성 완료: {result}')


if __name__ == '__main__':
    main()
