#!/usr/bin/env python3
"""Windows + 本机 Visio：对 build_docx.py 的输出做 COM 验证与收尾（可选步骤，失败不影响交付）。

用法：
  python visio_finish.py 输入.docx 输出.docx [--report report.json]
                         [--verify-only] [--no-emf] [--keep-xml-vsdx] [--move-mm 10]

对 docx 里每个嵌入的 vsdx，依次做：
  1. 打开（不可见 Visio 实例，独立进程，结束必退出）；
  2. 缓存值核对：读 Visio 算出的端点坐标，与 XML 里写的 V 值比对（偏差 > 0.05 mm 记为问题）；
  3. 粘附核对：每条 1-D 连线的 Connects 数量与 XML 里的 <Connect> 一致；
  4. 拖动测试：逐个移动有连线的形状 --move-mm，检查粘附的端点同步移动，再移回；
  5. 强制重算：把各连线 Begin/End 公式原样重写一遍，触发依赖重算；
  6. 导出 EMF（导出前临时加一个页面大小的透明框，保证 EMF 图幅 = 页面，导出后删掉）；
  7. 用 Visio 另存 vsdx（--keep-xml-vsdx 时保留套件生成的原文件）；
  8. 回写 docx：嵌入 vsdx 换成 Visio 存的版本，PNG 预览换成 EMF，补 emf 内容类型。
--verify-only：只做 1~4，不改任何文件，只出报告。

依赖：Windows、桌面版 Visio、pip install pywin32。docx 回写部分只用标准库。
"""
import argparse, math, json, os, re, shutil, sys, tempfile, zipfile

TOL_MM = 0.05
VS_NS = '{http://schemas.microsoft.com/office/visio/2012/main}'


# ----------------------------------------------------------------- docx 部分（纯标准库，可跨平台测试）
def _rb(path):
    with open(path, 'rb') as fh:
        return fh.read()


def list_embeds(docx):
    """返回 [(嵌入路径 word/embeddings/x.vsdx, 预览路径 word/media/y.*), ...]，按正文顺序。"""
    with zipfile.ZipFile(docx) as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf8')
        doc = z.read('word/document.xml').decode('utf8')
    rid = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
           for r in re.findall(r'<Relationship [^>]*?/>', rels)}
    out = []
    for m in re.finditer(r'<w:object [^>]*>.*?</w:object>', doc, flags=re.S):
        obj = m.group(0)
        emb = [rid[x] for x in re.findall(r'r:id="(rId\d+)"', obj) if rid.get(x, '').lower().endswith('.vsdx')]
        img = re.search(r'<v:imagedata r:id="(rId\d+)"', obj)
        if emb and img:
            out.append(('word/' + emb[0], 'word/' + rid[img.group(1)]))
    return out


def rewrite_docx(src, dst, replace_vsdx, replace_preview):
    """replace_vsdx: {zip内vsdx路径: 新文件}；replace_preview: {zip内旧预览路径: 新EMF文件}。"""
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    rels_p = 'word/_rels/document.xml.rels'
    rels = data[rels_p].decode('utf8')
    for old, emf in replace_preview.items():
        new = os.path.splitext(old)[0] + '.emf'
        rels = rels.replace(f'Target="{old[len("word/"):]}"', f'Target="{new[len("word/"):]}"')
        if old != new:
            del data[old]
            names[names.index(old)] = new
        data[new] = _rb(emf)
    data[rels_p] = rels.encode('utf8')
    for p, f in replace_vsdx.items():
        data[p] = _rb(f)
    ct = data['[Content_Types].xml'].decode('utf8')
    if replace_preview and 'Extension="emf"' not in ct:
        ct = ct.replace('<Default Extension="rels"', '<Default Extension="emf" ContentType="image/x-emf"/><Default Extension="rels"', 1)
        data['[Content_Types].xml'] = ct.encode('utf8')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])


def xml_expect(vsdx):
    """从套件生成的 vsdx 读出：{shape_id: {'BeginX':mm,...,'glued':{'BeginX','EndX'}}}（仅 1-D 连线）。"""
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(vsdx) as z:
        root = ET.fromstring(z.read('visio/pages/page1.xml'))
    exp = {}
    for sh in root.iter(VS_NS + 'Shape'):
        cells = {c.get('N'): c.get('V') for c in sh.findall(VS_NS + 'Cell')}
        if 'BeginX' in cells and 'EndX' in cells:
            exp[int(sh.get('ID'))] = {k: float(cells[k]) * 25.4 for k in ('BeginX', 'BeginY', 'EndX', 'EndY')}
            exp[int(sh.get('ID'))]['glued'] = set()
    for c in root.iter(VS_NS + 'Connect'):
        fid = int(c.get('FromSheet'))
        if fid in exp:
            exp[fid]['glued'].add(c.get('FromCell'))
    return exp


def emf_aspect(path):
    data = _rb(path)[:200]
    l, t, r, b = (int.from_bytes(data[24 + 4 * k:28 + 4 * k], 'little', signed=True) for k in range(4))
    return (r - l) / float(b - t) if b > t else 0.0


# ----------------------------------------------------------------- Visio COM 部分（仅 Windows）
def mm(cell):
    return cell.ResultIU * 25.4


def geom_issues(s, sid, tag):
    """连线几何（Geometry1）首点应在 Begin、末点应在 End；只看 Begin/End 单元格测不出线身脱节。"""
    out = []
    try:
        n = s.RowCount(10) - 1                     # visSectionFirstComponent，第 0 行是组件行
        if n < 2:
            return out
        bx, by = mm(s.CellsU('BeginX')), mm(s.CellsU('BeginY'))
        ex, ey = mm(s.CellsU('EndX')), mm(s.CellsU('EndY'))
        lx, ly = mm(s.CellsU('LocPinX')), mm(s.CellsU('LocPinY'))
        px, py = mm(s.CellsU('PinX')), mm(s.CellsU('PinY'))
        to_page = lambda x, y: (px - lx + x, py - ly + y)    # 1-D 连线 Angle=0
        g0 = to_page(mm(s.CellsU('Geometry1.X1')), mm(s.CellsU('Geometry1.Y1')))
        g1 = to_page(mm(s.CellsU(f'Geometry1.X{n}')), mm(s.CellsU(f'Geometry1.Y{n}')))
        for (gx, gy), (cx, cy), which in ((g0, (bx, by), '首点≠Begin'), (g1, (ex, ey), '末点≠End')):
            d = math.hypot(gx - cx, gy - cy)
            if d > TOL_MM:
                out.append(f'{tag}连线 {sid} 几何{which}（偏 {d:.2f} mm）')
    except Exception as e:                          # 单元格不存在等
        out.append(f'{tag}连线 {sid} 几何读取失败：{e}')
    return out


def check_drawing(app, path, exp, move_mm, verify_only, out_vsdx, out_emf):
    rep = {'issues': [], 'connectors': len(exp), 'glued_ends': sum(len(e['glued']) for e in exp.values())}
    doc = app.Documents.OpenEx(path, 0x8 + 0x80)       # visOpenDontList | visOpenMacrosDisabled
    try:
        pg = doc.Pages.Item(1)
        shapes = {pg.Shapes.Item(i).ID: pg.Shapes.Item(i) for i in range(1, pg.Shapes.Count + 1)}
        # 2. 缓存值核对 + 3. 粘附核对
        for sid, e in exp.items():
            s = shapes.get(sid)
            if s is None:
                rep['issues'].append(f'连线 {sid} 在 Visio 中不存在')
                continue
            for k in ('BeginX', 'BeginY', 'EndX', 'EndY'):
                d = abs(mm(s.CellsU(k)) - e[k])
                if d > TOL_MM:
                    rep['issues'].append(f'连线 {sid} {k} 缓存值偏差 {d:.2f} mm')
            rep['issues'] += geom_issues(s, sid, '')
            n = s.Connects.Count
            if n != len(e['glued']):
                rep['issues'].append(f'连线 {sid} 粘附端点 Visio={n} XML={len(e["glued"])}')
        # 4. 拖动测试：移动被粘附的 2-D 形状，检查粘附端跟随
        targets = {}
        for sid, s in shapes.items():
            if sid in exp:
                for j in range(1, s.Connects.Count + 1):
                    c = s.Connects.Item(j)
                    targets.setdefault(c.ToSheet.ID, []).append((sid, 'Begin' if 'Begin' in c.FromCell.Name else 'End'))
        rep['drag_tested'] = len(targets)
        for tid, ends in targets.items():
            if tid in exp or tid not in shapes:      # 只拖 2-D 形状
                continue
            t = shapes[tid]
            before = {(sid, end): (mm(shapes[sid].CellsU(end + 'X')), mm(shapes[sid].CellsU(end + 'Y'))) for sid, end in ends}
            pin = t.CellsU('PinX').FormulaU
            t.CellsU('PinX').ResultIU = t.CellsU('PinX').ResultIU + move_mm / 25.4
            for (sid, end), (x0, y0) in before.items():
                x1, y1 = mm(shapes[sid].CellsU(end + 'X')), mm(shapes[sid].CellsU(end + 'Y'))
                if abs(x1 - x0 - move_mm) > TOL_MM or abs(y1 - y0) > TOL_MM:
                    rep['issues'].append(f'拖动形状 {tid}：连线 {sid} 的 {end} 端未跟随（Δx={x1-x0:.2f} mm）')
            for sid in {sid for sid, _ in ends}:
                rep['issues'] += geom_issues(shapes[sid], sid, f'拖动形状 {tid} 后：')
            t.CellsU('PinX').FormulaU = pin
        if verify_only:
            return rep
        # 5. 强制重算
        for sid in exp:
            s = shapes.get(sid)
            if s is not None:
                for k in ('BeginX', 'BeginY', 'EndX', 'EndY'):
                    s.CellsU(k).FormulaForceU = s.CellsU(k).FormulaU
        # 6. 导出 EMF（临时透明页框，固定图幅）
        if out_emf:
            W, H = pg.PageSheet.CellsU('PageWidth').ResultIU, pg.PageSheet.CellsU('PageHeight').ResultIU
            fr = pg.DrawRectangle(0, 0, W, H)
            fr.CellsU('LinePattern').FormulaU = '0'
            fr.CellsU('FillPattern').FormulaU = '0'
            pg.Export(out_emf)
            fr.Delete()
            a, pa = emf_aspect(out_emf), W / H
            if abs(a - pa) / pa > 0.01:
                rep['issues'].append(f'EMF 纵横比 {a:.3f} 与页面 {pa:.3f} 不符，保留 PNG 预览')
                os.remove(out_emf)
        # 7. 另存
        if out_vsdx:
            doc.SaveAs(out_vsdx)
        return rep
    finally:
        doc.Saved = True
        doc.Close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('out', nargs='?')
    ap.add_argument('--report'); ap.add_argument('--verify-only', action='store_true')
    ap.add_argument('--no-emf', action='store_true'); ap.add_argument('--keep-xml-vsdx', action='store_true')
    ap.add_argument('--move-mm', type=float, default=10.0)
    a = ap.parse_args()
    if not a.verify_only and not a.out:
        sys.exit('需要输出路径，或加 --verify-only')
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    work = tempfile.mkdtemp(prefix='visio_finish_')
    app = win32com.client.DispatchEx('Visio.InvisibleApp')   # 独立实例，不接管用户已打开的 Visio
    app.AlertResponse = 1
    report, rep_vsdx, rep_prev = {}, {}, {}
    try:
        embeds = list_embeds(a.src)
        with zipfile.ZipFile(a.src) as z:
            for i, (emb, prev) in enumerate(embeds, 1):
                name = os.path.basename(emb)
                src = os.path.join(work, name)
                with open(src, 'wb') as fh:
                    fh.write(z.read(emb))
                out_v = None if (a.verify_only or a.keep_xml_vsdx) else os.path.join(work, 'out_' + name)
                out_e = None if (a.verify_only or a.no_emf) else os.path.join(work, name[:-5] + '.emf')
                try:
                    r = check_drawing(app, src, xml_expect(src), a.move_mm, a.verify_only, out_v, out_e)
                except Exception as ex:                       # 单图失败不影响其他图
                    r = {'issues': [f'Visio 处理失败：{ex}'], 'failed': True}
                    out_v = out_e = None
                report[name] = r
                if out_v and os.path.exists(out_v):
                    rep_vsdx[emb] = out_v
                if out_e and os.path.exists(out_e):
                    rep_prev[prev] = out_e
                print(f'#{i:<3}{name:<36} 连线 {r.get("connectors", "-")}  问题 {len(r["issues"])}')
                for s in r['issues']:
                    print('     ', s)
        if not a.verify_only:
            rewrite_docx(a.src, a.out, rep_vsdx, rep_prev)
            print('written', a.out, f'（vsdx 替换 {len(rep_vsdx)}，EMF 预览 {len(rep_prev)}）')
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        shutil.rmtree(work, ignore_errors=True)
    if a.report:
        with open(a.report, 'w', encoding='utf8') as fh:
            json.dump(report, fh, ensure_ascii=False, indent=1)
    sys.exit(1 if any(r['issues'] for r in report.values()) else 0)


if __name__ == '__main__':
    main()
