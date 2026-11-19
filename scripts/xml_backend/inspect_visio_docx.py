#!/usr/bin/env python3
"""通用诊断：列出 docx 中每个嵌入 Visio 对象的尺寸、缩放、字号与连线问题。

用法：python3 inspect_visio_docx.py <file.docx>
只读，不修改文件。依赖：Python 3 标准库。
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}


def rels_map(z, path):
    m = {}
    for r in re.finditer(r'<Relationship [^>]*?/>', z.read(path).decode('utf8')):
        s = r.group(0)
        m[re.search(r'Id="([^"]+)"', s).group(1)] = re.search(r'Target="([^"]+)"', s).group(1)
    return m


def cell(sh, name):
    c = sh.find(f'v:Cell[@N="{name}"]', NS)
    return c if c is not None else None


def inspect_vsdx(blob):
    import io
    vz = zipfile.ZipFile(io.BytesIO(blob))
    pages = vz.read('visio/pages/pages.xml').decode('utf8')
    pw = float(re.search(r'N="PageWidth" V="([^"]+)"', pages).group(1)) * 25.4
    ph = float(re.search(r'N="PageHeight" V="([^"]+)"', pages).group(1)) * 25.4
    page_files = sorted(n for n in vz.namelist() if re.match(r'visio/pages/page\d+\.xml$', n))
    stats = dict(pages=len(page_files), shapes=0, lines=0, slanted=0, unglued=0, bare=0, value_err=0, sizes=set())
    for pf in page_files:
        raw = vz.read(pf).decode('utf8')
        stats['value_err'] += raw.count('#VALUE!') + raw.count('#REF!') + raw.count('#NAME?')
        root = ET.fromstring(raw)
        glued = {}
        for c in root.iter(f'{{{NS["v"]}}}Connect'):
            glued.setdefault(c.get('FromSheet'), set()).add(c.get('FromCell'))
        for sh in root.iter(f'{{{NS["v"]}}}Shape'):
            stats['shapes'] += 1
            for ch in sh.iter(f'{{{NS["v"]}}}Cell'):
                if ch.get('N') == 'Size' and ch.get('V'):
                    try:
                        stats['sizes'].add(round(float(ch.get('V')) * 72, 1))
                    except ValueError:
                        pass
            bx = cell(sh, 'BeginX')
            if bx is None:
                continue
            stats['lines'] += 1
            if glued.get(sh.get('ID'), set()) != {'BeginX', 'EndX'}:
                # 无箭头又一端都没粘的线多是图例线、刻度线，单独计数，不算问题
                arrows = any(cell(sh, k) is not None and cell(sh, k).get('V') not in (None, '0') for k in ('BeginArrow', 'EndArrow'))
                if not arrows and sh.get('ID') not in glued:
                    stats['bare'] += 1
                else:
                    stats['unglued'] += 1
            # slanted segment check on local geometry
            pts = []
            for row in sh.iter(f'{{{NS["v"]}}}Row'):
                if row.get('T') in ('MoveTo', 'LineTo'):
                    xs = row.find('v:Cell[@N="X"]', NS)
                    ys = row.find('v:Cell[@N="Y"]', NS)
                    if xs is not None and ys is not None:
                        try:
                            pts.append((float(xs.get('V')), float(ys.get('V'))))
                        except (TypeError, ValueError):
                            pass
            ang = cell(sh, 'Angle')
            a = float(ang.get('V')) if ang is not None and ang.get('V') else 0.0
            import math
            ca, sa = math.cos(a), math.sin(a)
            gp = [(x * ca - y * sa, x * sa + y * ca) for x, y in pts]
            for (x1, y1), (x2, y2) in zip(gp, gp[1:]):
                if abs(x2 - x1) > 0.01 and abs(y2 - y1) > 0.01:  # > 0.25 mm 偏移
                    stats['slanted'] += 1
                    break
    return pw, ph, stats


def main(path):
    z = zipfile.ZipFile(path)
    doc = z.read('word/document.xml').decode('utf8')
    rels = rels_map(z, 'word/_rels/document.xml.rels')
    objs = re.findall(r'<w:object [^>]*>.*?</w:object>', doc, flags=re.S)
    print(f'{len(objs)} 个嵌入对象')
    for i, obj in enumerate(objs, 1):
        prog = re.search(r'ProgID="([^"]+)"', obj)
        ole = re.search(r'<o:OLEObject [^>]*r:id="([^"]+)"', obj)
        img = re.search(r'<v:imagedata r:id="([^"]+)"', obj)
        st = re.search(r'style="width:([\d.]+)pt;height:([\d.]+)pt', obj)
        lock = 'aspectratio="t"' in obj.split('<v:shape ', 1)[-1].split('</v:shape>')[0]
        target = rels.get(ole.group(1), '?') if ole else '?'
        line = f'#{i:02d} {prog.group(1) if prog else "?"} {target.split("/")[-1]} 预览={rels.get(img.group(1), "?").split("/")[-1] if img else "?"}'
        if not target.endswith('.vsdx'):
            print(line, '(非 vsdx，跳过)')
            continue
        pw, ph, s = inspect_vsdx(z.read('word/' + target))
        ww = float(st.group(1)) / 72 * 25.4 if st else 0
        wh = float(st.group(2)) / 72 * 25.4 if st else 0
        sx, sy = (ww / pw if pw else 0), (wh / ph if ph else 0)
        flags = []
        if abs(sx - sy) > 0.02:
            flags.append(f'非等比拉伸 {sx/sy-1:+.0%}')
        if abs(sx - 1) > 0.02:
            flags.append(f'Word 中缩放 {sx:.0%}')
        if not lock:
            flags.append('未锁定纵横比')
        if s['value_err']:
            flags.append(f'公式错误 {s["value_err"]} 处')
        if s['slanted']:
            flags.append(f'斜线 {s["slanted"]} 条')
        if s['unglued']:
            flags.append(f'未粘附连线 {s["unglued"]} 条')
        if s['bare']:
            line += f'  （另有无箭头未粘附线 {s["bare"]} 条，多为图例/刻度线）'
        eff = sorted(round(x * sx, 1) for x in s['sizes'])
        if eff and eff[0] < 7.5:
            flags.append(f'最小有效字号 {eff[0]}pt')
        print(line)
        print(f'    页面 {pw:.1f}x{ph:.1f}mm  Word {ww:.1f}x{wh:.1f}mm  形状 {s["shapes"]} 连线 {s["lines"]}  '
              f'字号(页内) {sorted(s["sizes"])}')
        print('    问题：' + ('；'.join(flags) if flags else '无'))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
