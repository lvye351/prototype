#!/usr/bin/env python3
"""从 PPTX 提取幻灯片 blocks（文本 + 表格）写入 assets/ppt-henan-ai-deck.json"""
import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def cell_text(tc):
    parts = []
    for t in tc.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            parts.append(t.text)
        if t.tail:
            parts.append(t.tail)
    return ''.join(parts).strip()


def parse_table(tbl_el):
    rows = []
    for tr in tbl_el.findall('a:tr', NS):
        row = [cell_text(tc) for tc in tr.findall('a:tc', NS)]
        if row:
            rows.append(row)
    return rows


def shape_text(sp):
    paras = []
    for ap in sp.findall('.//a:p', NS):
        parts = []
        for t in ap.findall('.//a:t', NS):
            if t.text:
                parts.append(t.text)
        s = ''.join(parts).strip()
        if s:
            paras.append(s)
    return paras


def collect_from_container(container, blocks):
    for child in container:
        tag = child.tag.split('}')[-1]
        if tag == 'sp':
            paras = shape_text(child)
            if paras:
                blocks.append({'type': 'text', 'paragraphs': paras})
        elif tag == 'graphicFrame':
            tbl = child.find('.//a:tbl', NS)
            if tbl is not None:
                rows = parse_table(tbl)
                if rows:
                    blocks.append({'type': 'table', 'rows': rows})
        elif tag == 'grpSp':
            collect_from_container(child, blocks)


def parse_slide(xml_bytes, idx):
    root = ET.fromstring(xml_bytes)
    blocks = []
    sp_tree = root.find('.//p:spTree', NS)
    if sp_tree is not None:
        collect_from_container(sp_tree, blocks)
    paras = []
    for b in blocks:
        if b['type'] == 'text':
            paras.extend(b['paragraphs'])
    title = paras[0] if paras else f'幻灯片 {idx}'
    return {
        'index': idx,
        'title': title[:120],
        'blocks': blocks,
        'paragraphs': paras,
        'body': '\n'.join(paras[1:8]) if len(paras) > 1 else '',
    }


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else (
        '/Users/wanghailong/部门/AI/文档生成助手/河南电信大模型二期/范文/'
        '01 中国电信河南公司2025年AI+服务管理能力升级研发项目公开询比采购方案汇报-打分标准maas.pptx'
    )
    out = sys.argv[2] if len(sys.argv) > 2 else 'assets/ppt-henan-ai-deck.json'
    z = zipfile.ZipFile(src)
    names = sorted(
        [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)],
        key=lambda x: int(re.search(r'slide(\d+)', x).group(1)),
    )
    slides = [parse_slide(z.read(n), i + 1) for i, n in enumerate(names)]
    deck = {
        'version': 2,
        'source': src.split('/')[-1],
        'title': '中国电信河南公司2025年AI+服务管理能力升级研发项目公开询比采购方案汇报',
        'slideCount': len(slides),
        'slides': slides,
    }
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(deck, f, ensure_ascii=False, indent=2)
    tables = sum(1 for s in slides for b in s['blocks'] if b['type'] == 'table')
    print(f'OK: {len(slides)} slides, {tables} tables -> {out}')


if __name__ == '__main__':
    main()
