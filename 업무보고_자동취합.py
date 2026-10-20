"""Windows 한/글 설치 필요. pip install pyhwpx lxml pywin32"""
import copy
import calendar
import datetime as dt
from dataclasses import dataclass
import queue
import re
import sys
import json
import tempfile
import threading
import traceback
import uuid
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
from lxml import etree as E

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
NS = {'hp': HP}
KINDS = ('추진실적', '추진계획')
# 기본 화면과 주요 강조는 무채색, 보조 선택 표시만 차분한 파스텔 블루.
ACCENT_BLUE = '#dce6f0'
TEMPLATE_PATH = Path(__file__).resolve().parent / 'assets' / 'report_template.hwpx'
# 업무 제목: 1. 제목, 1.1. 제목, 4.1. 제목. 날짜/소수점은 제외.
PREFIX = re.compile(r'^\s*(?:\d{1,3}\s*\.\s*)+(?!\d)')
ITEM_TOKEN = re.compile(r'^\s*REPORT_ITEM_[0-9a-f]{32}\s*')


def report_date(value):
    value = value.strip().rstrip('.')
    parts = re.fullmatch(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', value)
    if not parts:
        raise ValueError('날짜를 2026.09.10. 또는 2026-09-10 형식으로 입력하세요.')
    return dt.date(*map(int, parts.groups()))


def output_name(date, extension='.hwp'):
    return f'생방기술국_회의자료_{report_date(date):%y%m%d}{extension}'


def weekday_name(date):
    return '월화수목금토일'[report_date(date).weekday()]


@dataclass(frozen=True)
class WorkItem:
    key: str
    kind: str
    title: str
    source: Path
    package: Path
    first: int
    end: int


def reorder(items, source, destination):
    if 0 <= source < len(items) and 0 <= destination < len(items):
        items.insert(destination, items.pop(source))


def local(e):
    return E.QName(e).localname


def section_config(e):
    return local(e) == 'secPr' or (local(e) == 'ctrl' and any(
        local(x) in {'colPr', 'header', 'footer', 'pageNum'} for x in e.iter()))


def texts(p):
    # 표/그림/글상자 내부의 문자열은 업무 제목 판정에 사용하지 않는다.
    return p.findall('./hp:run/hp:t', NS)


def text(p):
    return ''.join(''.join(t.itertext()) for t in texts(p))


def heading(p):
    s = re.sub(r'\s+', '', text(p))
    s = s.strip('□■▣○●◆◇▶▷※[]【】<>〈〉:：.-0123456789')
    return s if s in KINDS else None


def blank(p):
    # 개체/필드/구역/탭/강제 줄바꿈이 있으면 절대로 빈 문단 취급하지 않는다.
    if p.get('pageBreak', '0') != '0' or p.get('columnBreak', '0') != '0':
        return False
    return not text(p).strip() and all(
        local(x) in {'p', 'run', 't', 'linesegarray', 'lineseg'}
        for x in p.iter()
    )


class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.data = {n: z.read(n) for n in z.namelist()}
        sections = [n for n in self.data
                    if re.fullmatch(r'Contents/section\d+\.xml', n)]
        if len(sections) != 1:
            raise ValueError(f'{Path(path).name}: 다중 구역 문서는 지원하지 않습니다. '
                             '구역별 파일로 나누어 등록하세요.')
        self.section_name = sections[0]
        parser = E.XMLParser(resolve_entities=False, no_network=True)
        self.root = E.fromstring(self.data[self.section_name], parser)
        self.header = E.fromstring(self.data['Contents/header.xml'], parser)
        self.paras = self.root.findall('hp:p', NS)
        self.shapes = {x.get('id'): x for x in self.header.iter()
                       if local(x) == 'paraPr'}
        self.next_shape = max(map(int, self.shapes), default=-1) + 1
        self.cache = {}

    def numbered(self, p):
        shape = self.shapes.get(p.get('paraPrIDRef'))
        if shape is None:
            raise ValueError('문단 모양 참조를 찾지 못했습니다.')
        h = next((x for x in shape if local(x) == 'heading'), None)
        return h is not None and h.get('type') == 'NUMBER' and h.get('level', '0') == '0'

    def no_number(self, p):
        old = p.get('paraPrIDRef')
        if old not in self.cache:
            original = self.shapes.get(old)
            if original is None:
                raise ValueError('문단 모양 참조를 찾지 못했습니다.')
            new = copy.deepcopy(original)
            ident = str(self.next_shape)
            self.next_shape += 1
            new.set('id', ident)
            for x in new:
                if local(x) == 'heading':
                    x.set('type', 'NONE')
                    x.set('idRef', '0')
                    x.set('level', '0')
            original.getparent().append(new)
            parent = original.getparent()
            parent.set('itemCnt', str(len(parent)))
            self.shapes[ident] = new
            self.cache[old] = ident
        p.set('paraPrIDRef', self.cache[old])

    def simple(self, template, value):
        p = E.Element(f'{{{HP}}}p', dict(template.attrib))
        p.set('id', '0')
        p.set('pageBreak', '0')
        p.set('columnBreak', '0')
        run0 = template.find('hp:run', NS)
        run = E.SubElement(p, f'{{{HP}}}run',
                           charPrIDRef=run0.get('charPrIDRef', '0') if run0 is not None else '0')
        E.SubElement(run, f'{{{HP}}}t').text = value
        self.no_number(p)
        return p

    def renumber(self, p, number):
        ts = texts(p)
        if any(len(t) for t in ts):
            raise ValueError('업무 제목에 탭/필드/강제 줄바꿈이 있습니다. 제목을 일반 텍스트로 바꿔주세요.')
        s = text(p)
        match = PREFIX.match(s)
        cut = match.end() if match else len(s) - len(s.lstrip())
        # 제목이 1FM처럼 숫자로 시작해도 번호 뒤 공백을 정확히 하나로 만든다.
        tail = s[cut:]
        cut += len(tail) - len(tail.lstrip())
        # 번호 부분만 제거. 나머지 글자 모양과 개체를 유지한다.
        remaining = cut
        for t in ts:
            value = t.text or ''
            take = min(len(value), remaining)
            t.text = value[take:]
            remaining -= take
        if not ts:
            raise ValueError('자동 번호 제목에 편집 가능한 텍스트가 없습니다.')
        ts[0].text = f'{number}. ' + (ts[0].text or '')
        self.no_number(p)
        for x in list(p):
            if local(x) == 'linesegarray':
                p.remove(x)  # 한/글에서 변경된 문단의 줄 배치를 다시 계산

    def items(self, source, kinds=KINDS):
        """표/그림을 포함한 업무 전체의 본문 문단 범위를 기록한다."""
        positions = [(i, heading(p)) for i, p in enumerate(self.paras) if heading(p)]
        if len({k for _, k in positions}) != len(positions):
            raise ValueError(f'{source.name}: 구분 제목이 중복됩니다.')
        result = {k: [] for k in KINDS}
        for n, (i, kind) in enumerate(positions):
            if kind not in kinds:
                continue
            end = positions[n + 1][0] if n + 1 < len(positions) else len(self.paras)
            nonempty = [j for j in range(i + 1, end) if not blank(self.paras[j])]
            if not nonempty:
                continue
            starts = [j for j in nonempty if PREFIX.match(text(self.paras[j]))
                      or self.numbered(self.paras[j])]
            if not starts:
                # 번호가 전혀 없는 파일은 첫 제목의 문단/글자 모양으로 후보를 찾는다.
                starts = [nonempty[0]]
            signatures = {self.title_signature(self.paras[j]) for j in starts}
            for j in nonempty:
                p = self.paras[j]
                s = text(p).strip()
                if (j not in starts and s and not s.startswith(('-', '※', '○', '●', '·', '▶'))
                        and self.title_signature(p) in signatures):
                    starts.append(j)
            starts.sort()
            if starts[0] != nonempty[0]:
                raise ValueError(f'{source.name}: {kind} 첫 업무 제목을 식별하지 못했습니다.')
            for pos, first in enumerate(starts):
                stop = starts[pos + 1] if pos + 1 < len(starts) else end
                title = PREFIX.sub('', text(self.paras[first]), count=1).strip()
                if not title:
                    raise ValueError(f'{source.name}: 업무 제목이 비어 있습니다.')
                result[kind].append(WorkItem(uuid.uuid4().hex, kind, title, source,
                                             Path(), first, stop))
        return result

    @staticmethod
    def title_signature(p):
        return (p.get('paraPrIDRef'), p.get('styleIDRef'),
                tuple(r.get('charPrIDRef') for r in p.findall('hp:run', NS)))

    def item_fragment(self, item, destination):
        out = [copy.deepcopy(p) for p in self.paras[item.first:item.end] if not blank(p)]
        if not out:
            raise ValueError('선택한 업무 내용이 비어 있습니다.')
        self.renumber(out[0], 1)
        # 최종 저장 직전에 이 표식으로 제목을 찾고 확정 순서대로 번호를 매긴다.
        ts = texts(out[0])
        ts[0].text = f'REPORT_ITEM_{item.key} ' + (ts[0].text or '')
        for p in out:
            for run in p.findall('hp:run', NS):
                for child in list(run):
                    if section_config(child):
                        run.remove(child)
        # HWPX 조각 자체는 유효한 구역 설정을 갖되 삽입할 때 기준 양식을 유지한다.
        sec = next((x for x in self.root.iter() if local(x) == 'secPr'), None)
        if sec is not None:
            out[0].find('hp:run', NS).insert(0, copy.deepcopy(sec))
        # InsertFile은 삽입 뒤 커서를 끝으로 옮기지 않는다.
        # 독립된 마지막 문단의 표식으로 삽입 끝을 정확히 찾는다.
        out.append(self.simple(out[0], f'REPORT_END_{item.key}'))
        root = copy.deepcopy(self.root)
        for child in list(root):
            root.remove(child)
        for p in out:
            root.append(p)
        data = dict(self.data)
        data[self.section_name] = E.tostring(root, encoding='UTF-8', xml_declaration=True)
        data['Contents/header.xml'] = E.tostring(self.header, encoding='UTF-8', xml_declaration=True)
        write_package(data, destination)


def write_package(data, destination):
    with zipfile.ZipFile(destination, 'w') as z:
        if 'mimetype' in data:
            z.writestr('mimetype', data['mimetype'], compress_type=zipfile.ZIP_STORED)
        for name, value in data.items():
            if name != 'mimetype':
                z.writestr(name, value, compress_type=zipfile.ZIP_DEFLATED)


def replace_span(p, start, end, replacement):
    """문자열의 특정 범위만 교체하고 주변 글자 모양은 유지."""
    offset = 0
    inserted = False
    for t in texts(p):
        if len(t):
            raise ValueError('날짜 문단에 특수 제어문자가 있습니다.')
        value = t.text or ''
        left, right = max(start - offset, 0), min(end - offset, len(value))
        if left < right:
            t.text = value[:left] + (replacement if not inserted else '') + value[right:]
            inserted = True
        offset += len(value)
    if not inserted:
        raise ValueError('날짜 교체 위치를 찾지 못했습니다.')
    for child in list(p):
        if local(child) == 'linesegarray':
            p.remove(child)


def template_skeleton(source, destination, date, weekday):
    pack = Package(source)
    # 첫 번째 파일에 한쪽 항목만 있어도 빠진 구분 제목을 사본에 만든다.
    present = [(p, heading(p)) for p in pack.paras if heading(p)]
    if len(present) == 1:
        old, kind = present[0]
        missing = '추진계획' if kind == '추진실적' else '추진실적'
        added = pack.simple(old, f'<{missing}>')
        gap = pack.simple(old, '')
        if missing == '추진실적':
            index = pack.root.index(old)
            pack.root.insert(index, added)
            pack.root.insert(index + 1, gap)
        else:
            pack.root.append(gap)
            pack.root.append(added)
        pack.paras = pack.root.findall('hp:p', NS)
    positions = [(i, heading(p)) for i, p in enumerate(pack.paras) if heading(p)]
    if [k for _, k in positions] != list(KINDS):
        raise ValueError('내장 양식의 추진실적/추진계획 구분 제목을 확인할 수 없습니다.')
    # 실제 양식: 2026.09.23.(수) 생방기술2. 날짜/요일 외 문자열은 유지.
    pattern = re.compile(r'(?P<date>\d{4}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*\d{1,2}\s*\.?)'
                         r'(?P<gap>\s*)\((?P<day>[월화수목금토일](?:요일)?)\)')
    found = []
    for p in pack.paras[:positions[0][0]]:
        for match in pattern.finditer(text(p)):
            found.append((p, match))
    if len(found) != 1:
        raise ValueError('내장 양식의 상단에서 날짜(요일)를 하나로 식별하지 못했습니다. '
                         '예: 2026.09.23.(수)')
    p, match = found[0]
    replace_span(p, match.start('day'), match.end('day'), weekday)
    replace_span(p, match.start('date'), match.end('date'), date)
    markers = {}
    gaps = {}
    for n in range(len(positions) - 1, -1, -1):
        i, kind = positions[n]
        end = positions[n + 1][0] if n + 1 < len(positions) else len(pack.paras)
        body = pack.paras[i + 1:end]
        trailing = []
        if n + 1 < len(positions):
            for candidate in reversed(body):
                if not blank(candidate):
                    break
                trailing.insert(0, copy.deepcopy(candidate))
        gaps[kind] = len(trailing)
        token = 'REPORT_BODY_' + uuid.uuid4().hex
        markers[kind] = token
        marker = pack.simple(body[0] if body else pack.paras[i], token)
        for old in body:
            pack.root.remove(old)
        insertion = pack.root.index(pack.paras[i]) + 1
        pack.root.insert(insertion, marker)
        for gap in trailing:
            insertion += 1
            pack.root.insert(insertion, gap)
    data = dict(pack.data)
    data[pack.section_name] = E.tostring(pack.root, encoding='UTF-8', xml_declaration=True)
    data['Contents/header.xml'] = E.tostring(pack.header, encoding='UTF-8', xml_declaration=True)
    write_package(data, destination)
    return markers, gaps


def body_spacing(source, destination, gaps, expected=None):
    """제목/날짜/상단 도형과 제목 앞 여백은 건드리지 않고 본문만 정리."""
    pack = Package(source)
    positions = [(i, heading(p)) for i, p in enumerate(pack.paras) if heading(p)]
    if [k for _, k in positions] != list(KINDS):
        raise ValueError('취합 후 양식의 구분 제목이 변경되었습니다.')
    for n in range(len(positions) - 1, -1, -1):
        i, kind = positions[n]
        end = positions[n + 1][0] if n + 1 < len(positions) else len(pack.paras)
        body = pack.paras[i + 1:end]
        kept = [p for p in body if not blank(p)]
        normalized = []
        keys = []
        for p in kept:
            match = ITEM_TOKEN.match(text(p))
            if expected is not None:
                if match:
                    key = match.group().strip().removeprefix('REPORT_ITEM_')
                    keys.append(key)
                    if len(keys) > 1:
                        normalized.append(pack.simple(p, ''))
                    replace_span(p, 0, match.end(), '')
                    pack.renumber(p, len(keys))
            elif PREFIX.match(text(p)) and normalized:
                normalized.append(pack.simple(p, ''))
            normalized.append(p)
        if expected is not None and keys != [item.key for item in expected[kind]]:
            raise ValueError(f'{kind}: 목록 {len(expected[kind])}개 / 확인된 제목 {len(keys)}개. '
                             '업무 순서 또는 개수가 일치하지 않아 저장을 중단합니다.')
        if any('REPORT_END_' in text(p) for p in normalized):
            raise ValueError(f'{kind}: 업무 끝 표식이 남았습니다. 저장을 중단합니다.')
        for _ in range(gaps[kind]):
            normalized.append(pack.simple(pack.paras[i], ''))
        for old in body:
            pack.root.remove(old)
        insertion = pack.root.index(pack.paras[i]) + 1
        for p in normalized:
            pack.root.insert(insertion, p)
            insertion += 1
    data = dict(pack.data)
    data[pack.section_name] = E.tostring(pack.root, encoding='UTF-8', xml_declaration=True)
    data['Contents/header.xml'] = E.tostring(pack.header, encoding='UTF-8', xml_declaration=True)
    write_package(data, destination)


def extract_previous_plans(source, folder, log):
    """지난주 최종본의 생방기술2 추진계획만 독립된 사본으로 추출."""
    with zipfile.ZipFile(source) as z:
        data = {name: z.read(name) for name in z.namelist()}
    names = sorted((n for n in data if re.fullmatch(r'Contents/section\d+\.xml', n)),
                   key=lambda n: int(re.search(r'section(\d+)', n).group(1)))
    if not names:
        raise ValueError('지난주 최종본에서 문서 구역을 읽을 수 없습니다.')
    parser = E.XMLParser(resolve_entities=False, no_network=True)
    roots = [E.fromstring(data[n], parser) for n in names]
    # 부서별 제목이 없는 최종본: 추진계획 안의 마지막 생방기술1 업무를 경계로 사용한다.
    # 같은 제목이 추진실적에도 있으므로 문서 전체에서 단순 검색하지 않는다.
    paras = [p for r in roots for p in r.findall('hp:p', NS)]
    ranges = []
    for i, p in enumerate(paras):
        if heading(p) == '추진계획':
            end = next((j for j in range(i + 1, len(paras)) if heading(paras[j])), len(paras))
            ranges.append((i, end))
    boundaries = []
    for first, end in ranges:
        for i in range(first + 1, end):
            title = PREFIX.sub('', text(paras[i])).strip()
            if re.sub(r'\s+', '', title) == '견학및기타제작사항':
                boundaries.append((first, end, i))
    if len(boundaries) != 1:
        raise ValueError('지난주 추진계획에서 ‘견학 및 기타 제작사항’ 업무를 '
                         f'하나로 찾지 못했습니다 ({len(boundaries)}개). 생방기술1이 섞이지 않도록 중단합니다.')
    first, end, boundary = boundaries[0]
    header = E.fromstring(data['Contents/header.xml'], parser)
    shapes = {x.get('id'): x for x in header.iter() if local(x) == 'paraPr'}

    def is_title(p):
        # 표 내부 문단은 후보에 포함하지 않는다. 직접 입력/자동 번호 모두 지원한다.
        if PREFIX.match(text(p)):
            return True
        shape = shapes.get(p.get('paraPrIDRef'))
        if shape is None or not text(p).strip():
            return False
        return any(local(x) == 'heading' and x.get('type') == 'NUMBER'
                   and x.get('level') == '0' for x in shape)

    start = next((i for i in range(boundary + 1, end) if is_title(paras[i])), None)
    if start is None:
        raise ValueError('지난주 추진계획의 경계 업무 다음에서 생방기술2 업무 제목을 찾지 못했습니다.')
    plan_heading = paras[first]
    selected = [copy.deepcopy(p) for p in paras[start:end]]
    log('지난주 추진계획의 ‘견학 및 기타 제작사항’ 본문·표 다음 업무부터 가져옵니다: '
        + text(paras[start]).strip())
    root = copy.deepcopy(roots[0])
    for child in list(root):
        root.remove(child)
    out = [copy.deepcopy(plan_heading)] + selected
    for p in out:
        for run in p.findall('hp:run', NS):
            for child in list(run):
                if section_config(child):
                    run.remove(child)
    sec = next((x for r in roots for x in r.iter() if local(x) == 'secPr'), None)
    if sec is not None:
        run = out[0].find('hp:run', NS)
        if run is None:
            run = E.SubElement(out[0], f'{{{HP}}}run', charPrIDRef='0')
        run.insert(0, copy.deepcopy(sec))
    for p in out:
        root.append(p)
    for name in names[1:]:
        data.pop(name)
    data[names[0]] = E.tostring(root, encoding='UTF-8', xml_declaration=True)
    header = E.fromstring(data['Contents/header.xml'], parser)
    header.set('secCnt', '1')
    data['Contents/header.xml'] = E.tostring(header, encoding='UTF-8', xml_declaration=True)
    if 'Contents/content.hpf' in data:
        manifest = E.fromstring(data['Contents/content.hpf'], parser)
        removed_ids = set()
        for node in list(manifest.iter()):
            if local(node) == 'item' and node.get('href') in names[1:]:
                removed_ids.add(node.get('id'))
                node.getparent().remove(node)
        for node in list(manifest.iter()):
            if local(node) == 'itemref' and node.get('idref') in removed_ids:
                node.getparent().remove(node)
        data['Contents/content.hpf'] = E.tostring(manifest, encoding='UTF-8', xml_declaration=True)
    target = Path(folder) / 'previous_department2_plans.hwpx'
    write_package(data, target)
    log('생방기술2 추진계획만 추출했습니다. 생방기술1 및 지난주 추진실적은 제외합니다.')
    return target


def prepare(paths, folder, log, previous_path=None):
    """1단계: 파일을 임시 사본으로 변환하고 업무 제목을 불러온다."""
    import pythoncom
    from pyhwpx import Hwp
    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = Hwp(new=True, visible=True)
        items = {kind: [] for kind in KINDS}
        if previous_path is None:
            raise ValueError('지난주 최종본을 선택하세요.')
        previous_path = Path(previous_path)
        previous_target = folder / 'previous.hwpx'
        log(f'지난주 최종본 읽는 중: {previous_path.name}')
        if not hwp.open(str(previous_path)) or not hwp.save_as(str(previous_target), format='HWPX'):
            raise RuntimeError('지난주 최종본 열기/변환 실패')
        hwp.clear(1)
        previous_pack_path = extract_previous_plans(previous_target, folder, log)
        previous_pack = Package(previous_pack_path)
        previous_items = previous_pack.items(previous_path, kinds=('추진계획',))['추진계획']
        if not previous_items:
            raise ValueError('지난주 생방기술2 추진계획에서 업무를 찾지 못했습니다.')
        items['추진실적'] = [WorkItem(item.key, '추진실적', item.title, previous_path,
                                     previous_pack_path, item.first, item.end) for item in previous_items]
        log(f'이번 주 추진실적: 지난주 생방기술2 추진계획 {len(previous_items)}개 (지난주 순서로 불러옴 · 이후 순서 조정 가능)')
        for n, path in enumerate(paths):
            log(f'제목 읽는 중: {path.name}')
            target = folder / f'source_{n}.hwpx'
            if not hwp.open(str(path)) or not hwp.save_as(str(target), format='HWPX'):
                raise RuntimeError(f'파일 읽기/변환 실패: {path.name}')
            hwp.clear(1)
            pack = Package(target)
            extracted = pack.items(path, kinds=('추진계획',))
            for kind in ('추진계획',):
                for item in extracted[kind]:
                    items[kind].append(WorkItem(item.key, item.kind, item.title,
                                      item.source, target, item.first, item.end))
                log(f'{path.name}: {kind} {len(extracted[kind])}개')
        if not items['추진계획']:
            raise ValueError('이번 주 취합 파일에서 추진계획 업무를 찾지 못했습니다.')
        return items
    finally:
        if hwp is not None:
            try:
                hwp.clear(1)
                hwp.quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def review_styles(source, destination, log):
    """최종 문단 스타일/들여쓰기 검사. 표 안 문단과 개체 배치는 유지."""
    pack = Package(source)
    styles = {x.get('name'): x for x in pack.header.iter()
              if local(x) == 'style' and x.get('type') == 'PARA'}
    # 원본 양식의 중점 스타일 이름은 특수 글꼴 문자로 저장된 경우가 있다.
    if '·' not in styles:
        alias = styles.get('\uf0a0')
        if alias is None:
            for p in pack.paras:
                if text(p).lstrip().startswith('·'):
                    alias = next((s for s in styles.values()
                                  if s.get('id') == p.get('styleIDRef')), None)
                    if alias is not None:
                        break
        if alias is None:
            raise ValueError('내장 양식에서 · 문단의 스타일을 찾지 못했습니다.')
        cloned = copy.deepcopy(alias)
        ident = str(max(int(s.get('id')) for s in styles.values()) + 1)
        cloned.set('id', ident)
        cloned.set('name', '·')
        cloned.set('engName', '')
        cloned.set('nextStyleIDRef', ident)
        parent = alias.getparent()
        parent.append(cloned)
        parent.set('itemCnt', str(len(parent)))
        styles['·'] = cloned
        log('양식의 중점 들여쓰기를 유지한 · 스타일을 결과 사본에 추가했습니다.')
    required = ('바탕글', '날짜', '추진', '1,2,3', '·', '-')
    missing = [name for name in required if name not in styles]
    if missing:
        raise ValueError('내장 양식의 필수 스타일이 없습니다: ' + ', '.join(missing))
    first_heading = next(i for i,p in enumerate(pack.paras) if heading(p))
    counts = {name: 0 for name in required}
    checked = []
    title_shapes = {}

    def apply(p, name, title=False, preserve_layout=False):
        style = styles[name]
        p.set('styleIDRef', style.get('id'))
        target = p.get('paraPrIDRef') if preserve_layout else style.get('paraPrIDRef')
        if title:
            # 상단 제목의 가운데 정렬/줄 간격은 유지하고 들여쓰기만 바탕글에 맞춘다.
            old = p.get('paraPrIDRef')
            if old not in title_shapes:
                cloned = copy.deepcopy(pack.shapes[old])
                base = pack.shapes[target]
                for tag in ('intent', 'left', 'right'):
                    a = [x for x in cloned.iter() if local(x) == tag]
                    b = [x for x in base.iter() if local(x) == tag]
                    for x,y in zip(a,b):
                        x.attrib.clear()
                        x.attrib.update(y.attrib)
                ident = str(pack.next_shape)
                pack.next_shape += 1
                cloned.set('id', ident)
                parent = pack.shapes[old].getparent()
                parent.append(cloned)
                parent.set('itemCnt', str(len(parent)))
                pack.shapes[ident] = cloned
                title_shapes[old] = ident
            target = title_shapes[old]
        p.set('paraPrIDRef', target)
        # 자동 번호/글머리표와 수동으로 적힌 기호가 중복되지 않게 한다.
        pack.no_number(p)
        for child in list(p):
            if local(child) == 'linesegarray':
                p.remove(child)
        counts[name] += 1
        checked.append((p, name))

    for i,p in enumerate(pack.paras):
        s = text(p)
        if i < first_heading:
            if re.match(r'^\s*\d{4}[.\-/]', s):
                apply(p, '날짜')
            elif s.strip():
                apply(p, '바탕글', title=True)
            # 이 양식의 상단 제목은 도형 안에 들어 있다.
            for q in p.findall('./hp:run/hp:container//hp:p', NS):
                if text(q).strip() and not any(local(a) == 'tbl' for a in q.iterancestors()):
                    apply(q, '바탕글', title=True)
            continue
        if heading(p):
            apply(p, '추진')
        elif re.match(r'^\s*(\d{1,3})\.', s):
            number = int(re.match(r'^\s*(\d{1,3})\.', s).group(1))
            pack.renumber(p, number)
            apply(p, '1,2,3')
        elif s.lstrip().startswith(('·', '-')):
            marker = s.lstrip()[0]
            # Shift+Tab으로 지정한 내어쓰기/탭/왼쪽 여백은 직접 문단 서식에 저장된다.
            # 스타일 이름만 맞추고 원본 문단 서식과 앞쪽 공백·탭은 유지한다.
            apply(p, marker, preserve_layout=True)
    for p,name in checked:
        if p.get('styleIDRef') != styles[name].get('id'):
            raise ValueError(f'{name} 스타일 적용 검사 실패')
        if name == '1,2,3' and not re.match(r'^\d+\. \S', text(p)):
            raise ValueError('번호 형식 검사 실패: 숫자 + 점 + 공백 한 칸 + 제목')
    data = dict(pack.data)
    data[pack.section_name] = E.tostring(pack.root, encoding='UTF-8', xml_declaration=True)
    data['Contents/header.xml'] = E.tostring(pack.header, encoding='UTF-8', xml_declaration=True)
    write_package(data, destination)
    log('최종 스타일 검사 완료 (본문의 원본 내어쓰기·탭 정렬 유지): ' + ' / '.join(f'{k} {v}문단' for k,v in counts.items()))


def merge(ordered, output, date, weekday, template, log):
    import pythoncom
    from pyhwpx import Hwp
    pythoncom.CoInitialize()
    hwp = None
    temp_dir = None
    try:
        weekday = weekday_name(date)
        hwp = Hwp(new=True, visible=True)
        temp_dir = tempfile.TemporaryDirectory(prefix='report_merge_', ignore_cleanup_errors=True)
        with temp_dir as folder:
            folder = Path(folder)
            if not hwp.open(str(template)):
                raise RuntimeError('내장 양식 열기 실패')
            template_copy = folder / 'template.hwpx'
            if not hwp.save_as(str(template_copy), format='HWPX'):
                raise RuntimeError('양식 변환 실패')
            hwp.clear(1)
            skeleton = folder / 'skeleton.hwpx'
            markers, gaps = template_skeleton(template_copy, skeleton, date, weekday)
            fragments = {kind: [] for kind in KINDS}
            for kind in KINDS:
                for n, item in enumerate(ordered[kind]):
                    pack = Package(item.package)
                    frag = folder / f'{kind}_{n}.hwpx'
                    pack.item_fragment(item, frag)
                    fragments[kind].append(frag)
                    log(f'{kind} 최종 {n + 1}. {item.title}')
            if not any(fragments.values()):
                raise ValueError('취합할 업무를 찾지 못했습니다.')
            if not hwp.open(str(skeleton)):
                raise RuntimeError('양식 사본 열기 실패')
            for kind in reversed(KINDS):
                hwp.MoveDocBegin()
                if not hwp.find(markers[kind], direction='AllDoc', AutoSpell=0):
                    raise RuntimeError(f'{kind} 삽입 위치를 찾지 못했습니다.')
                hwp.Delete()
                for n, frag in enumerate(fragments[kind]):
                    # 양식의 구역/용지/머리말을 유지하고 내용 서식만 가져온다.
                    if not hwp.insert_file(str(frag), keep_section=0,
                                           keep_charshape=1, keep_parashape=1, keep_style=0):
                        raise RuntimeError(f'삽입 실패: {frag.name}')
                    # 문서 전체의 끝이 아니라 방금 넣은 업무의 끝으로 이동한다.
                    # 표식을 지우면 독립된 빈 문단에 커서가 남아 다음 제목이 붙지 않는다.
                    end_token = f'REPORT_END_{ordered[kind][n].key}'
                    hwp.MoveDocBegin()
                    if not hwp.find(end_token, direction='AllDoc', AutoSpell=0):
                        raise RuntimeError(f'{kind} {n + 1}번 업무의 삽입 끝을 찾지 못했습니다.')
                    hwp.Delete()
            # 본문 안에서만 공백 정리: 양식 상단의 빈 문단은 보존한다.
            combined = folder / 'combined.hwpx'
            if not hwp.save_as(str(combined), format='HWPX'):
                raise RuntimeError('중간 저장 실패')
            hwp.clear(1)  # 검사/예외 처리 전에 combined.hwpx의 잠금을 해제한다.
            cleaned = folder / 'cleaned.hwpx'
            body_spacing(combined, cleaned, gaps, ordered)
            reviewed = folder / 'reviewed.hwpx'
            review_styles(cleaned, reviewed, log)
            hwp.clear(1)
            if not hwp.open(str(reviewed)):
                raise RuntimeError('스타일 정리 결과 열기 실패')
            suffix = output.suffix.lower()
            if not hwp.save_as(str(output), format='HWPX' if suffix == '.hwpx' else 'HWP'):
                raise RuntimeError('결과 저장 실패')
            hwp.clear(1)  # 임시 파일 잠금 해제
            if not hwp.open(str(output)):
                raise RuntimeError('저장했지만 결과 다시 열기에 실패했습니다.')
            log(f'저장 완료: {output}')
    except Exception:
        if hwp is not None:
            try:
                hwp.clear(1)
                hwp.quit()
            except Exception:
                pass
        raise
    finally:
        if temp_dir is not None:
            try:
                temp_dir.cleanup()  # 오류 시 한/글을 닫은 다음 다시 정리한다.
            except OSError:
                log(f'임시 파일 정리를 건너뛰었습니다: {temp_dir.name}')
        pythoncom.CoUninitialize()


def enable_dpi_awareness():
    """PY 실행에서도 Tk 창을 만들기 전에 Windows의 흐릿한 비트맵 확대를 해제."""
    if sys.platform != 'win32':
        return
    import ctypes
    user32 = ctypes.windll.user32
    try:
        user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetProcessDpiAwarenessContext.restype = ctypes.c_bool
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
        # EXE는 매니페스트로 이미 설정되어 있으므로 다시 설정할 필요가 없다.
        if ctypes.windll.kernel32.GetLastError() == 5:
            return
    except AttributeError:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        user32.SetProcessDPIAware()


def scale_widget_spacing(parent, scale):
    """글자는 포인트로 선명하게 그리고, 여백은 현재 화면 배율에 맞춘다."""
    def scaled(value):
        parts = parent.tk.splitlist(value) if isinstance(value, str) else value
        if not isinstance(parts, (tuple, list)):
            parts = (parts,)
        return tuple(round(float(x)*scale) for x in parts)

    for child in parent.winfo_children():
        manager = child.winfo_manager()
        if manager in ('pack', 'grid'):
            info = child.pack_info() if manager == 'pack' else child.grid_info()
            padding = {key: scaled(info[key]) for key in ('padx', 'pady', 'ipadx', 'ipady') if key in info}
            if padding:
                if manager == 'pack':
                    child.pack_configure(**padding)
                else:
                    child.grid_configure(**padding)
        if isinstance(child, ttk.Frame):
            value = child.cget('padding')
            if value:
                child.configure(padding=scaled(value))
        scale_widget_spacing(child, scale)



class Card(tk.Frame):
    """창 크기에 맞춰 늘어나는 둥근 흰색 카드."""
    def __init__(self, parent):
        super().__init__(parent, bg='#f2f2f2', bd=0)
        self.canvas = tk.Canvas(self, bg='#f2f2f2', highlightthickness=0, bd=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.body = ttk.Frame(self, style='Card.TFrame', padding=3)
        self.body.pack(fill='both', expand=True, padx=14, pady=14)
        self.tk.call('lower', self.canvas._w)
        self.bind('<Configure>', self.paint)

    def paint(self, event):
        w, h, r = event.width, event.height, round(18*self.winfo_fpixels('1i')/96)
        self.canvas.delete('card')
        self.canvas.create_polygon(
            r, 0, w-r, 0, w, 0, w, r, w, h-r, w, h,
            w-r, h, r, h, 0, h, 0, h-r, 0, r, 0, 0,
            smooth=True, splinesteps=24, fill='white', outline='', tags='card')


class DatePicker(tk.Toplevel):
    def __init__(self, owner, selected, on_select):
        super().__init__(owner)
        self.withdraw()
        self.title('회의자료 날짜 선택')
        self.transient(owner)
        self.resizable(False, False)
        self.configure(bg='white')
        self.selected = selected
        self.year, self.month = selected.year, selected.month
        self.on_select = on_select
        self.days = {}
        self.ui_scale = owner.ui_scale
        body = ttk.Frame(self, padding=14, style='Card.TFrame')
        body.pack(fill='both', expand=True)
        nav = ttk.Frame(body)
        nav.pack(fill='x', pady=(0, 12))
        ttk.Button(nav, text='‹', width=3, command=lambda: self.change_month(-1)).pack(side='left')
        self.caption = ttk.Label(nav, style='Heading.TLabel', anchor='center')
        self.caption.pack(side='left', fill='x', expand=True, padx=12)
        ttk.Button(nav, text='›', width=3, command=lambda: self.change_month(1)).pack(side='right')
        self.grid_days = ttk.Frame(body)
        self.grid_days.pack()
        footer = ttk.Frame(body)
        footer.pack(fill='x', pady=(12, 0))
        ttk.Button(footer, text='오늘', command=lambda: self.select(dt.date.today())).pack(side='left')
        ttk.Button(footer, text='닫기', command=self.destroy).pack(side='right')
        self.bind('<Escape>', lambda e: self.destroy())
        scale_widget_spacing(body, self.ui_scale)
        self.render()
        self.update_idletasks()
        anchor = owner.date_entry
        x = min(anchor.winfo_rootx(), self.winfo_screenwidth()-self.winfo_reqwidth()-16)
        y = anchor.winfo_rooty()+anchor.winfo_height()+6
        if y+self.winfo_reqheight() > self.winfo_screenheight()-50:
            y = anchor.winfo_rooty()-self.winfo_reqheight()-6
        self.geometry(f'+{max(0, x)}+{max(0, y)}')
        self.deiconify()
        # 날짜 입력칸은 달력이 열린 동안에도 직접 입력할 수 있다.
        owner.date_entry.after_idle(owner.date_entry.focus_set)

    def change_month(self, delta):
        total = (self.year-1)*12+self.month-1+delta
        if not 0 <= total < 9999*12:
            return
        self.year, month = divmod(total, 12)
        self.year += 1
        self.month = month+1
        self.render()

    def render(self):
        self.caption.config(text=f'{self.year}년 {self.month}월')
        for widget in self.grid_days.winfo_children():
            widget.destroy()
        self.days.clear()
        for col, label in enumerate('월화수목금토일'):
            color = '#525252'
            ttk.Label(self.grid_days, text=label, anchor='center', foreground=color).grid(row=0, column=col, pady=(0, round(6*self.ui_scale)))
        for row, week in enumerate(calendar.Calendar(firstweekday=0).monthdayscalendar(self.year, self.month), 1):
            for col, day in enumerate(week):
                if not day:
                    continue
                value = dt.date(self.year, self.month, day)
                button = ttk.Button(self.grid_days, text=str(day), width=3,
                                    style='SelectedDay.TButton' if value == self.selected else 'Day.TButton',
                                    command=lambda v=value: self.select(v))
                button.grid(row=row, column=col, padx=round(2*self.ui_scale), pady=round(2*self.ui_scale))
                self.days[day] = button

    def select(self, value):
        self.on_select(value)
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        enable_dpi_awareness()
        super().__init__()
        self.ui_scale = max(1.0, self.winfo_fpixels('1i')/96)
        self.title('주간 회의자료 취합')
        self.configure(bg='#f2f2f2')
        self.setup_theme()
        width = min(round(1180*self.ui_scale), self.winfo_screenwidth()-round(80*self.ui_scale))
        height = min(round(850*self.ui_scale), self.winfo_screenheight()-round(100*self.ui_scale))
        self.geometry(f'{width}x{height}')
        self.minsize(min(round(930*self.ui_scale), width), min(round(720*self.ui_scale), height))
        self.paths = []
        self.previous_path = None
        self.loaded_previous = None
        self.items = {kind: [] for kind in KINDS}
        self.item_lists = {}
        self.item_counts = {}
        self.drag_start = {}
        self.cache_dir = None
        self.loaded_paths = None
        self.events = queue.Queue()
        self.busy = False
        self.calendar_popup = None
        self.date = tk.StringVar(value=dt.date.today().strftime('%Y.%m.%d.'))
        self.weekday = tk.StringVar(value=weekday_name(self.date.get())+'요일')
        self.output = tk.StringVar()
        self.previous_text = tk.StringVar(value='선택된 파일이 없습니다')
        self.preview_state = tk.StringVar(value='파일 선택 후 업무 제목을 불러오세요.')

        box = ttk.Frame(self, style='Page.TFrame', padding=20)
        box.pack(fill='both', expand=True)
        box.columnconfigure(0, weight=0, minsize=round(326*self.ui_scale))
        box.columnconfigure(1, weight=1)
        box.rowconfigure(1, weight=1)
        header = ttk.Frame(box, style='Page.TFrame')
        header.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 16))
        ttk.Label(header, text='주간 회의자료', style='Title.TLabel').pack(side='left')
        ttk.Label(header, text='생방기술2', style='Badge.TLabel', padding=(12, 6)).pack(side='right')

        left_shell = ttk.Frame(box, style='Page.TFrame')
        left_shell.grid(row=1, column=0, sticky='nsew', padx=(0, 14))
        left_shell.rowconfigure(0, weight=1)
        left_shell.columnconfigure(0, weight=1)
        self.left_canvas = tk.Canvas(left_shell, width=round(326*self.ui_scale), height=1, bg='#f2f2f2', bd=0, highlightthickness=0)
        self.left_canvas.grid(row=0, column=0, sticky='nsew')
        left_scroll = ttk.Scrollbar(left_shell, orient='vertical', command=self.left_canvas.yview)
        left_scroll.grid(row=0, column=1, sticky='ns')
        self.left_canvas.config(yscrollcommand=left_scroll.set)
        left = ttk.Frame(self.left_canvas, style='Page.TFrame')
        left_window = self.left_canvas.create_window((0, 0), window=left, anchor='nw')
        def resize_left(event=None):
            self.left_canvas.itemconfigure(left_window, width=self.left_canvas.winfo_width(),
                                           height=max(left.winfo_reqheight(), self.left_canvas.winfo_height()))
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox('all'))
        left.bind('<Configure>', resize_left)
        self.left_canvas.bind('<Configure>', resize_left)
        def scroll_inputs(event):
            if isinstance(event.widget, tk.Listbox):
                return
            canvas = self.left_canvas
            if (canvas.winfo_rootx() <= event.x_root < canvas.winfo_rootx()+canvas.winfo_width()
                    and canvas.winfo_rooty() <= event.y_root < canvas.winfo_rooty()+canvas.winfo_height()):
                canvas.yview_scroll(-int(event.delta/120), 'units')
        self.bind('<MouseWheel>', scroll_inputs, add='+')
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        card = Card(left)
        card.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        body = card.body
        ttk.Label(body, text='01  지난주 최종본', style='Heading.TLabel').pack(anchor='w')
        ttk.Label(body, text='생방기술2의 계획을 이번 주 실적으로', style='Muted.TLabel').pack(anchor='w', pady=(3, 10))
        ttk.Entry(body, textvariable=self.previous_text, state='readonly', width=28).pack(fill='x')
        ttk.Button(body, text='최종본 선택', command=self.choose_previous).pack(fill='x', pady=(10, 0))

        card = Card(left)
        card.grid(row=1, column=0, sticky='nsew', pady=(0, 10))
        body = card.body
        ttk.Label(body, text='02  이번 주 취합 파일', style='Heading.TLabel').pack(anchor='w')
        ttk.Label(body, text='각 파일의 추진계획을 가져옵니다', style='Muted.TLabel').pack(anchor='w', pady=(3, 10))
        files = ttk.Frame(body)
        files.pack(fill='both', expand=True)
        self.listbox = self.make_list(files, height=4)
        self.listbox.config(width=24)
        file_buttons = ttk.Frame(body)
        file_buttons.pack(fill='x', pady=(8, 0))
        for label, command in [('추가', self.add), ('삭제', self.remove),
                               ('↑', lambda: self.move(-1)), ('↓', lambda: self.move(1))]:
            ttk.Button(file_buttons, text=label, width=4, command=command).pack(side='left', padx=(0, 5), expand=True, fill='x')
        self.load_button = ttk.Button(body, text='업무 제목 불러오기', style='Soft.TButton', command=self.load_titles)
        self.load_button.pack(fill='x', pady=(10, 0))

        card = Card(left)
        card.grid(row=2, column=0, sticky='ew')
        body = card.body
        ttk.Label(body, text='03  날짜와 저장 위치', style='Heading.TLabel').pack(anchor='w')
        date_row = ttk.Frame(body)
        date_row.pack(fill='x', pady=(12, 4))
        self.date_entry = ttk.Entry(date_row, textvariable=self.date, width=13)
        self.date_entry.pack(side='left', fill='x', expand=True)
        self.date_entry.bind('<Button-1>', lambda e: self.open_calendar())
        self.date_entry.bind('<Alt-Down>', lambda e: self.open_calendar())
        self.date_entry.bind('<Escape>', lambda e: self.close_calendar())
        self.calendar_button = ttk.Button(date_row, text='달력', width=4, command=self.open_calendar)
        self.calendar_button.pack(side='left', padx=(5, 0))
        ttk.Label(date_row, textvariable=self.weekday, style='Accent.TLabel').pack(side='left', padx=(8, 0))
        ttk.Label(body, text='직접 입력하거나 달력에서 선택하세요', style='Muted.TLabel').pack(anchor='w', pady=(0, 10))
        ttk.Entry(body, textvariable=self.output, state='readonly', width=28).pack(fill='x')
        ttk.Button(body, text='저장 폴더 선택', command=self.choose_output).pack(fill='x', pady=(8, 0))

        right = ttk.Frame(box, style='Page.TFrame')
        right.grid(row=1, column=1, sticky='nsew')
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)
        card = Card(right)
        card.grid(row=0, column=0, sticky='nsew', pady=(0, 10))
        body = card.body
        ttk.Label(body, text='업무 순서', style='Heading.TLabel').pack(anchor='w')
        ttk.Label(body, text='두 항목 모두 드래그 또는 이동 버튼으로 조정할 수 있습니다.', style='Muted.TLabel').pack(anchor='w', pady=(4, 6))
        self.preview_label = ttk.Label(body, textvariable=self.preview_state, style='Accent.TLabel', wraplength=round(500*self.ui_scale))
        self.preview_label.pack(anchor='w', fill='x', pady=(0, 12))
        body.bind('<Configure>', lambda e: self.preview_label.config(wraplength=max(round(180*self.ui_scale), e.width-round(38*self.ui_scale))))
        self.tabs = ttk.Notebook(body)
        self.tabs.pack(fill='both', expand=True)
        for kind in KINDS:
            panel = ttk.Frame(self.tabs, padding=(8, 12, 8, 4))
            self.tabs.add(panel, text=f'  {kind}  ')
            count = tk.StringVar(value='0개 업무')
            self.item_counts[kind] = count
            ttk.Label(panel, textvariable=count, style='Muted.TLabel').pack(anchor='w', pady=(0, 8))
            row_items = ttk.Frame(panel)
            row_items.pack(fill='both', expand=True)
            listing = self.make_list(row_items, height=8)
            listing.config(width=30)
            listing.bind('<ButtonPress-1>', lambda e, k=kind: self.drag_begin(k, e))
            listing.bind('<B1-Motion>', lambda e, k=kind: self.drag_motion(k, e))
            listing.bind('<ButtonRelease-1>', lambda e, k=kind: self.drag_end(k, e))
            listing.bind('<Alt-Up>', lambda e, k=kind: self.move_item(k, -1))
            listing.bind('<Alt-Down>', lambda e, k=kind: self.move_item(k, 1))
            self.item_lists[kind] = listing
            row_move = ttk.Frame(panel)
            row_move.pack(fill='x', pady=(10, 0))
            ttk.Button(row_move, text='↑ 위로', command=lambda k=kind: self.move_item(k, -1)).pack(side='left')
            ttk.Button(row_move, text='↓ 아래로', command=lambda k=kind: self.move_item(k, 1)).pack(side='left', padx=6)
        card = Card(right)
        card.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        body = card.body
        ttk.Label(body, text='진행 기록', style='Muted.TLabel').pack(anchor='w', pady=(0, 6))
        log_row = ttk.Frame(body)
        log_row.pack(fill='x')
        self.status = tk.Text(log_row, height=3, width=30, state='disabled', bd=0,
                              bg='white', fg='#404040', font=self.body_font, wrap='word', highlightthickness=0)
        self.status.pack(side='left', fill='x', expand=True)
        scroll = ttk.Scrollbar(log_row, orient='vertical', command=self.status.yview)
        scroll.pack(side='right', fill='y')
        self.status.config(yscrollcommand=scroll.set)
        self.start_button = ttk.Button(right, text='최종 회의자료 만들기  →', style='Primary.TButton', command=self.start)
        self.start_button.grid(row=2, column=0, sticky='ew')
        scale_widget_spacing(box, self.ui_scale)
        self.date.trace_add('write', self.update_filename)
        self.bind('<Button-1>', self.dismiss_calendar, add='+')
        self.after(100, self.poll)
        self.protocol('WM_DELETE_WINDOW', self.close)

    def setup_theme(self):
        # 글꼴은 이 프로세스에서만 사용하며 직원 PC의 시스템 글꼴을 바꾸지 않는다.
        if sys.platform == 'win32':
            import ctypes
            for path in (Path(__file__).resolve().parent/'assets'/'fonts').glob('*.otf'):
                ctypes.windll.gdi32.AddFontResourceExW(str(path), 0x10, 0)
        families = set(tkfont.families(self))
        heading_family = 'Pretendard' if 'Pretendard' in families else '맑은 고딕'
        self.body_family = 'Pretendard Medium' if 'Pretendard Medium' in families else heading_family
        self.body_font = (self.body_family, 10)
        for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont', 'TkCaptionFont'):
            tkfont.nametofont(name).configure(family=self.body_family, size=10)
        tkfont.nametofont('TkHeadingFont').configure(family=heading_family, size=10, weight='bold')
        p = lambda *values: tuple(round(v*self.ui_scale) for v in values)
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', font=self.body_font, background='white', foreground='#171717')
        style.configure('Page.TFrame', background='#f2f2f2')
        style.configure('Card.TFrame', background='white')
        style.configure('Title.TLabel', background='#f2f2f2', foreground='#111111', font=(heading_family, 21, 'bold'))
        style.configure('Heading.TLabel', foreground='#171717', font=(heading_family, 12, 'bold'))
        style.configure('Muted.TLabel', foreground='#525252')
        # 상태 설명은 검정/회색으로 읽기 쉽게 표시하고 블루는 선택·포커스에 사용.
        style.configure('Accent.TLabel', foreground='#333333')
        style.configure('Badge.TLabel', foreground='#171717', background='#e8e8e8', font=(heading_family, 10, 'bold'))
        style.configure('TButton', padding=p(12, 7), background='#eeeeee', foreground='#171717', borderwidth=0, relief='flat')
        style.map('TButton', background=[('disabled', '#eeeeee'), ('pressed', '#d9d9d9'), ('active', '#e2e2e2')], foreground=[('disabled', '#808080')])
        style.configure('Soft.TButton', background='#e8e8e8', foreground='#171717')
        # 주요 실행 버튼은 블랙 바탕·화이트 글씨로 뚜렷하게 강조한다.
        style.configure('Primary.TButton', padding=p(16, 13), background='#171717', foreground='white', font=(heading_family, 11, 'bold'))
        style.map('Primary.TButton', background=[('disabled', '#dddddd'), ('pressed', '#333333'), ('active', '#262626'), ('!disabled', '#171717')], foreground=[('disabled', '#737373'), ('!disabled', 'white')])
        style.configure('TEntry', padding=p(7), foreground='#171717', fieldbackground='#f7f7f7', bordercolor='#d9d9d9', lightcolor='#d9d9d9', darkcolor='#d9d9d9')
        style.map('TEntry', fieldbackground=[('readonly', '#f7f7f7')], bordercolor=[('focus', '#737373')])
        style.configure('TNotebook', background='white', borderwidth=0, bordercolor='white', lightcolor='white', darkcolor='white', tabmargins=(0, 0, 0, 0))
        style.configure('TNotebook.Tab', padding=p(14, 10), background='#eeeeee', foreground='#333333', borderwidth=0)
        style.map('TNotebook.Tab', padding=[('selected', p(14, 10)), ('!selected', p(14, 10))], background=[('selected', '#171717')], foreground=[('selected', 'white')], expand=[('selected', (0, 0, 0, 0)), ('!selected', (0, 0, 0, 0))])
        style.configure('TScrollbar', background='#cfcfcf', troughcolor='#f5f5f5', borderwidth=0, arrowsize=round(12*self.ui_scale))
        style.map('TScrollbar', background=[('active', '#bfbfbf'), ('pressed', '#b0b0b0')])
        style.configure('Day.TButton', padding=p(5, 7), background='white')
        style.configure('SelectedDay.TButton', padding=p(5, 7), background=ACCENT_BLUE, foreground='#111111')
        style.map('SelectedDay.TButton', background=[('!disabled', ACCENT_BLUE)], foreground=[('!disabled', '#111111')])
        self.option_add('*Listbox.font', self.body_font)
        self.option_add('*Listbox.background', '#ffffff')
        self.option_add('*Listbox.foreground', '#171717')
        self.option_add('*Listbox.selectBackground', ACCENT_BLUE)
        self.option_add('*Listbox.selectForeground', '#111111')


    def make_list(self, parent, height):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        listing = tk.Listbox(parent, height=height, exportselection=False, activestyle='none',
                             bd=0, highlightthickness=1, highlightbackground='#d9d9d9',
                             highlightcolor='#737373', selectborderwidth=0)
        vertical = ttk.Scrollbar(parent, orient='vertical', command=listing.yview)
        horizontal = ttk.Scrollbar(parent, orient='horizontal', command=listing.xview)
        listing.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        listing.grid(row=0, column=0, sticky='nsew')
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        return listing

    def open_calendar(self):
        if self.busy:
            return
        if self.calendar_popup is not None and self.calendar_popup.winfo_exists():
            self.calendar_popup.lift()
            return
        try:
            value = report_date(self.date.get())
        except ValueError:
            value = dt.date.today()
        self.calendar_popup = DatePicker(self, value, self.select_date)

    def select_date(self, value):
        self.date.set(value.strftime('%Y.%m.%d.'))
        self.date_entry.focus_set()

    def close_calendar(self):
        if self.calendar_popup is not None and self.calendar_popup.winfo_exists():
            self.calendar_popup.destroy()
        self.calendar_popup = None

    def dismiss_calendar(self, event):
        if event.widget not in (self.date_entry, self.calendar_button):
            self.close_calendar()


    def refresh(self, index=None):
        self.listbox.delete(0, 'end')
        for n, path in enumerate(self.paths, 1):
            self.listbox.insert('end', f'{n}. {path.name}')
        if index is not None:
            self.listbox.selection_set(index)
            self.listbox.see(index)

    def add(self):
        if self.busy:
            return
        changed = False
        for name in filedialog.askopenfilenames(filetypes=[('한글 문서', '*.hwp *.hwpx')]):
            path = Path(name).resolve()
            if path not in self.paths:
                self.paths.append(path)
                changed = True
        if changed:
            self.invalidate()
        self.refresh()

    def remove(self):
        if not self.busy and self.listbox.curselection():
            self.paths.pop(self.listbox.curselection()[0])
            self.invalidate()
            self.refresh()

    def move(self, delta):
        if self.busy or not self.listbox.curselection():
            return
        i = self.listbox.curselection()[0]
        j = i + delta
        if 0 <= j < len(self.paths):
            self.paths[i], self.paths[j] = self.paths[j], self.paths[i]
            self.invalidate()
            self.refresh(j)

    def choose_output(self):
        if self.busy:
            return
        name = filedialog.askdirectory(title='최종 회의자료를 저장할 폴더')
        if name:
            try:
                self.output.set(str(Path(name) / output_name(self.date.get())))
            except ValueError as e:
                messagebox.showerror('날짜 확인', str(e))

    def choose_previous(self):
        if self.busy:
            return
        name = filedialog.askopenfilename(title='지난주 최종 회의자료 선택', filetypes=[('한글 문서', '*.hwp *.hwpx')])
        if name:
            path = Path(name).resolve()
            if path != self.previous_path:
                self.previous_path = path
                self.previous_text.set(path.name)
                self.invalidate()

    def update_filename(self, *_):
        try:
            self.weekday.set(weekday_name(self.date.get()) + '요일')
        except ValueError:
            self.weekday.set('날짜 확인')
        if self.output.get():
            try:
                self.output.set(str(Path(self.output.get()).parent / output_name(self.date.get())))
            except ValueError:
                pass

    def invalidate(self):
        self.loaded_paths = None
        self.loaded_previous = None
        self.items = {kind: [] for kind in KINDS}
        for kind in KINDS:
            self.refresh_items(kind)
        self.preview_state.set('파일 목록이 바뀌었습니다. 업무 제목을 다시 불러오세요.')

    def refresh_items(self, kind, selected=None):
        listing = self.item_lists[kind]
        self.item_counts[kind].set(f'{len(self.items[kind])}개 업무')
        listing.delete(0, 'end')
        for n, item in enumerate(self.items[kind], 1):
            listing.insert('end', f'{n}. {item.title}    [{item.source.name}]')
        if selected is not None and self.items[kind]:
            listing.selection_set(selected)
            listing.see(selected)

    def drag_begin(self, kind, event):
        if not self.busy and self.items[kind]:
            self.drag_start[kind] = self.item_lists[kind].nearest(event.y)

    def drag_motion(self, kind, event):
        if self.busy or kind not in self.drag_start:
            return
        listing = self.item_lists[kind]
        if event.y < 10:
            listing.yview_scroll(-1, 'units')
        elif event.y > listing.winfo_height() - 10:
            listing.yview_scroll(1, 'units')
        target = listing.nearest(event.y)
        listing.selection_clear(0, 'end')
        listing.selection_set(target)
        listing.activate(target)

    def drag_end(self, kind, event):
        source = self.drag_start.pop(kind, None)
        if self.busy or source is None:
            return
        listing = self.item_lists[kind]
        if not (0 <= event.x < listing.winfo_width() and 0 <= event.y < listing.winfo_height()):
            self.refresh_items(kind, source)
            return
        target = listing.nearest(event.y)
        reorder(self.items[kind], source, target)
        self.refresh_items(kind, target)

    def move_item(self, kind, delta):
        if self.busy or not self.item_lists[kind].curselection():
            return
        source = self.item_lists[kind].curselection()[0]
        target = source + delta
        if 0 <= target < len(self.items[kind]):
            reorder(self.items[kind], source, target)
            self.refresh_items(kind, target)

    def load_titles(self):
        self.close_calendar()
        if self.busy:
            return
        if not self.paths:
            messagebox.showerror('파일 확인', '취합할 파일을 먼저 추가하세요.')
            return
        if self.previous_path is None or not self.previous_path.is_file():
            messagebox.showerror('지난주 최종본 확인', '지난주 최종본을 별도로 선택하세요.')
            return
        if self.previous_path in self.paths:
            messagebox.showerror('파일 확인', '지난주 최종본은 이번 주 취합 파일 목록에서 제외하세요.')
            return
        paths = list(self.paths)
        previous = self.previous_path
        cache = tempfile.TemporaryDirectory(prefix='report_preview_')
        self.busy = True
        self.load_button.config(state='disabled')
        self.start_button.config(state='disabled')
        self.preview_state.set('업무 제목을 불러오는 중입니다…')

        def worker():
            try:
                items = prepare(paths, Path(cache.name), lambda s: self.events.put(('log', s)), previous_path=previous)
                self.events.put(('loaded', (items, cache, paths, previous)))
            except Exception:
                error = traceback.format_exc()
                try:
                    cache.cleanup()
                except OSError:
                    pass
                self.events.put(('error', error))
        threading.Thread(target=worker, daemon=True).start()

    def start(self):
        self.close_calendar()
        if self.busy:
            return
        try:
            date = self.date.get().strip()
            report_date(date)
            weekday = weekday_name(date)
            if self.loaded_paths != self.paths or self.loaded_previous != self.previous_path:
                raise ValueError('먼저 업무 제목 불러오기를 실행한 뒤 순서를 조정하세요.')
            if self.cache_dir is None:
                raise ValueError('업무 제목을 먼저 불러오세요.')
            template = TEMPLATE_PATH
            if not template.is_file():
                raise ValueError('프로그램의 내장 자료를 읽을 수 없습니다. 배포 파일을 다시 받아주세요.')
            if not self.paths or not self.output.get().strip():
                raise ValueError('파일과 저장 위치를 지정하세요.')
            output = Path(self.output.get()).resolve()
            output = output.parent / output_name(date)
            self.output.set(str(output))
            ordered = {k: list(self.items[k]) for k in KINDS}
            if not all(ordered.values()):
                raise ValueError('추진실적과 추진계획을 모두 불러온 뒤 실행하세요.')
            if output.suffix.lower() not in ('.hwp', '.hwpx'):
                raise ValueError('결과 확장자는 .hwp 또는 .hwpx여야 합니다.')
            if output in self.paths or output == template or output == self.previous_path:
                raise ValueError('원본 파일을 결과 저장 위치로 지정할 수 없습니다.')
            if output.exists():
                raise ValueError('같은 이름의 결과 파일이 있습니다. 기존 결과의 이름을 바꾸거나 다른 저장 폴더를 선택하세요.')
            if not output.parent.is_dir():
                raise ValueError('저장 폴더가 존재하지 않습니다.')
        except ValueError as e:
            messagebox.showerror('입력 확인', str(e))
            return
        self.busy = True
        self.start_button.config(state='disabled')
        self.load_button.config(state='disabled')

        def worker():
            try:
                merge(ordered, output, date, weekday, template,
                      lambda s: self.events.put(('log', s)))
                self.events.put(('done', str(output)))
            except Exception:
                self.events.put(('error', traceback.format_exc()))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        while not self.events.empty():
            kind, value = self.events.get_nowait()
            if kind == 'loaded':
                items, cache, paths, previous = value
                if self.cache_dir is not None:
                    self.cache_dir.cleanup()
                self.cache_dir = cache
                self.items = items
                self.loaded_paths = paths
                self.loaded_previous = previous
                for section in KINDS:
                    self.refresh_items(section)
                self.preview_state.set(' / '.join(f'{k} {len(items[k])}개' for k in KINDS)
                                       + ' — 각 탭에서 제목 순서를 확인하고 드래그하세요.')
                self.busy = False
                self.start_button.config(state='normal')
                self.load_button.config(state='normal')
                continue
            self.status.config(state='normal')
            self.status.insert('end', value + '\n')
            self.status.see('end')
            self.status.config(state='disabled')
            if kind in ('done', 'error'):
                self.busy = False
                self.start_button.config(state='normal')
                self.load_button.config(state='normal')
                if kind == 'error':
                    self.preview_state.set('오류가 발생했습니다. 아래 내용을 확인하세요.')
                if kind == 'done':
                    messagebox.showinfo('취합 완료', value + '\n표·그림 배치를 확인하세요.')
                else:
                    messagebox.showerror('취합 실패', value)
        self.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo('처리 중', '취합이 끝난 뒤 창을 닫아주세요.')
        else:
            if self.cache_dir is not None:
                self.cache_dir.cleanup()
            self.destroy()


def bundle_self_test(source, output, report):
    """배포 파일의 실제 한/글 취합 경로를 점검하는 내부 검사."""
    logs = []
    result = {'ok': False}
    try:
        with tempfile.TemporaryDirectory(prefix='report_exe_test_', ignore_cleanup_errors=True) as folder:
            items = prepare([Path(source).resolve()], Path(folder), logs.append, previous_path=Path(source).resolve())
            ordered = {'추진실적': items['추진실적'], '추진계획': list(reversed(items['추진계획']))}
            merge(ordered, Path(output).resolve(), '2026.10.14.', '수', TEMPLATE_PATH, logs.append)
            result['ok'] = True
            result['items'] = {k: len(v) for k,v in ordered.items()}
    except Exception:
        result['error'] = traceback.format_exc()
    result['log'] = logs
    Path(report).resolve().write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == '__main__':
    if len(sys.argv) == 5 and sys.argv[1] == '--self-test':
        bundle_self_test(*sys.argv[2:])
    else:
        App().mainloop()
