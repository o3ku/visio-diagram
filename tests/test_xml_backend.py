"""XML 后端冒烟与回归测试（不需要 Visio）。

运行：python -m unittest discover -s tests -v
需要 Pillow 和一种中文字体（自动查找微软雅黑 / Noto Sans CJK / 苹方，或设 VISIO_KIT_FONT_R）。
"""
import contextlib, io, os, re, shutil, subprocess, sys, tempfile, unittest, zipfile
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XB = os.path.join(ROOT, 'scripts', 'xml_backend')
sys.path.insert(0, XB)

import vlib  # noqa: E402
import vsdxpkg  # noqa: E402
from vlib import Fig  # noqa: E402

NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
EXAMPLE = os.path.join(ROOT, 'assets', 'example_xml_flowchart.py')


def page_root(fig):
    return ET.fromstring(fig.visio_xml().split('\n', 1)[1])


def cells(sh):
    return {c.get('N'): c for c in sh.findall('v:Cell', NS)}


def connectors(root):
    return [sh for sh in root.iter(f'{{{NS["v"]}}}Shape') if 'BeginX' in cells(sh)]


def geom_rows(sh):
    g = sh.find('v:Section[@N="Geometry"]', NS)
    return [r for r in g.findall('v:Row', NS) if r.get('T') in ('MoveTo', 'LineTo')]


def eval_span(formula, W, H):
    """求值 vlib 写的几何公式：Width*k、Height*k、Width*k±n mm（结果单位英寸）。"""
    m = re.fullmatch(r'(Width|Height)\*([-\d.e]+)(?:([+-])([\d.]+) mm)?', formula)
    assert m, formula
    v = (W if m.group(1) == 'Width' else H) * float(m.group(2))
    if m.group(3):
        v += (1 if m.group(3) == '+' else -1) * float(m.group(4)) / 25.4
    return v


def example_fig():
    import importlib.util
    spec = importlib.util.spec_from_file_location('ex', EXAMPLE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m.FIGS[0][2]()


class TestVlib(unittest.TestCase):
    def test_example_glue(self):
        root = page_root(example_fig())
        lines = connectors(root)
        cons = root.findall('.//v:Connect', NS)
        self.assertEqual(len(cons), 2 * len(lines))
        ids = {sh.get('ID') for sh in root.iter(f'{{{NS["v"]}}}Shape')}
        for c in cons:
            self.assertIn(c.get('ToSheet'), ids)
            self.assertIn(c.get('FromPart'), ('9', '12'))
        for sh in lines:
            c = cells(sh)
            for k in ('BeginX', 'BeginY', 'EndX', 'EndY'):
                self.assertTrue(c[k].get('F', '').startswith('PAR(PNT('), k)

    def test_geometry_follows_any_end_move(self):
        """每个几何点都有公式；Width/Height 任意变化后首点在 (0,0)、末点在 (W,H)。含零宽竖线、零高横线。"""
        f = Fig(100, 60)
        f.box('a', 10, 5, 20, 10, 'A')
        f.box('b', 10, 40, 20, 10, 'B')
        f.box('c', 60, 5, 20, 10, 'C')
        f.conn('a', 'b', 'b', 't', label='是')                 # 竖线，Width=0
        f.conn('a', 'r', 'c', 'l')                              # 横线，Height=0
        f.line([(30, 10), (40, 10), (40, 45), (30, 45)])         # U 形回折，Width=0 且中间点有偏移
        root = page_root(f)
        for sh in connectors(root):
            rows = geom_rows(sh)
            for r in rows:
                for c in r.findall('v:Cell', NS):
                    self.assertTrue(c.get('F'), f'shape {sh.get("ID")} 几何行缺公式')
            for W, H in ((0.0, 1.0), (0.8, 1.0), (-0.5, 0.3), (1.2, 0.0)):
                x0 = eval_span(rows[0].find('v:Cell[@N="X"]', NS).get('F'), W, H)
                y0 = eval_span(rows[0].find('v:Cell[@N="Y"]', NS).get('F'), W, H)
                x1 = eval_span(rows[-1].find('v:Cell[@N="X"]', NS).get('F'), W, H)
                y1 = eval_span(rows[-1].find('v:Cell[@N="Y"]', NS).get('F'), W, H)
                self.assertAlmostEqual(x0, 0); self.assertAlmostEqual(y0, 0)
                self.assertAlmostEqual(x1, W); self.assertAlmostEqual(y1, H)
            c = cells(sh)
            if 'TxtPinX' in c:
                self.assertTrue(c['TxtPinX'].get('F')); self.assertTrue(c['TxtPinY'].get('F'))

    def test_u_shape_keeps_offset(self):
        f = Fig(100, 60)
        f.line([(30, 10), (40, 10), (40, 45), (30, 45)])
        rows = geom_rows(connectors(page_root(f))[0])
        xs = [r.find('v:Cell[@N="X"]', NS).get('F') for r in rows]
        self.assertEqual(xs[0], 'Width*0'); self.assertEqual(xs[-1], 'Width*1')
        self.assertEqual(xs[1], 'Width*0+10.0000 mm'); self.assertEqual(xs[2], 'Width*1+10.0000 mm')

    def test_band_not_glue_target(self):
        f = Fig(100, 60)
        f.container('grp', 10, 10, 80, 40, '分组', tw=30)
        f.line([(20, 0), (20, 10)])                             # 落在容器顶边、标题带范围内
        root = page_root(f)
        shapes = list(root.iter(f'{{{NS["v"]}}}Shape'))
        container_id, band_id = shapes[0].get('ID'), shapes[1].get('ID')
        to = [c.get('ToSheet') for c in root.findall('.//v:Connect', NS)]
        self.assertEqual(to, [container_id])
        self.assertNotIn(band_id, to)

    def test_align_right(self):
        f = Fig(50, 20)
        f.box('r', 5, 5, 40, 10, '右对齐', align='r')
        sh = next(root_sh for root_sh in page_root(f).iter(f'{{{NS["v"]}}}Shape'))
        ha = sh.find('.//v:Section[@N="Paragraph"]/v:Row/v:Cell[@N="HorzAlign"]', NS)
        self.assertEqual(ha.get('V'), '2')
        with self.assertRaises(ValueError):
            f.box('x', 0, 0, 1, 1, align='right')

    def test_errors(self):
        f = Fig(50, 20)
        f.box('a', 0, 0, 10, 10)
        with self.assertRaises(ValueError):
            f.box('a', 20, 0, 10, 10)
        with self.assertRaisesRegex(KeyError, 'nope'):
            f.conn('a', 'r', 'nope', 'l')
        with self.assertRaises(ValueError):
            f.line([(1, 1)], label='x')

    def test_png(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, 'x.png')
            example_fig().render_png(p, S=4)
            self.assertGreater(os.path.getsize(p), 1000)


def _lum(hexcolor):
    r, g, b = (int(hexcolor[i:i + 2], 16) / 255 for i in (1, 3, 5))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = f(r), f(g), f(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast(a, b):
    la, lb = _lum(a), _lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _themed_fig(theme):
    f = Fig(100, 60, theme=theme)
    f.container('grp', 5, 5, 90, 50, '分组')
    f.box('a', 10, 15, 30, 12, '节点', 'node')
    f.box('k', 55, 15, 30, 12, '重点', 'key')
    f.box('d', 10, 35, 30, 12, '关键', 'dark')
    f.box('o', 55, 35, 15, 12, '通过', 'ok')
    f.box('w', 72, 35, 15, 12, '异常', 'warn')
    f.conn('a', 'r', 'k', 'l', label='是')
    return f


class TestThemes(unittest.TestCase):
    def test_registered_in_both_routes(self):
        """COM 主题注册表与 XML 色板注册表同名同义，不缺不漏。"""
        import theme
        self.assertEqual(sorted(vlib.list_themes()), sorted(theme.THEMES))
        self.assertEqual(sorted(vlib.list_themes()),
                         ['mono', 'tech-cn'])
        for name in vlib.list_themes():
            theme.get(name)  # COM 路线能取到
            theme.xml_palette(name)  # XML 路线能取到

    def test_default_is_tech_cn_and_unchanged(self):
        f = Fig(50, 20)
        self.assertEqual(f.theme, 'tech-cn')
        self.assertEqual(vlib.STYLES['dark'], ('#1F3864', '#1F3864', 0, '#FFFFFF', 1, 9.0, 0.75))
        self.assertEqual(vlib.STYLES['node'][0], '#FFFFFF')
        with self.assertRaises(KeyError):
            Fig(50, 20, theme='nope')
        with self.assertRaises(KeyError):
            vlib.set_default_theme('nope')

    def test_each_theme_renders_glues_and_parses(self):
        for name in vlib.list_themes():
            f = _themed_fig(name)
            root = page_root(f)
            glued, ends = f.glue_report()
            self.assertEqual((glued, ends), (2, 2), name)
            with tempfile.TemporaryDirectory() as d:
                p = os.path.join(d, 'x.png')
                f.render_png(p, S=4)
                self.assertGreater(os.path.getsize(p), 1000, name)

    def test_theme_colors_in_xml(self):
        expect = {'tech-cn': '#1F3864', 'mono': '#000000'}
        for name, color in expect.items():
            xml = _themed_fig(name).visio_xml()
            self.assertIn(color, xml, name)
        # mono 产物是纯灰阶（除白底外所有填充/线条色 r==g==b）
        xml = _themed_fig('mono').visio_xml()
        for m in re.finditer(r'(?:FillForegnd|LineColor|Color)" V="(#[0-9A-Fa-f]{6})"', xml):
            h = m.group(1)
            self.assertEqual(h[1:3], h[3:5], (h, 'mono 混入彩色'))
            self.assertEqual(h[3:5], h[5:7], (h, 'mono 混入彩色'))
        self.assertNotIn('#1F3864', xml)

    def test_text_fill_contrast(self):
        """每套主题的文字/填充对比度 ≥ 4.5:1（含 8pt note）。"""
        for name in vlib.list_themes():
            styles = vlib._xml_palette(name)['styles']
            for role in ('node', 'key', 'dark', 'ok', 'warn', 'soft', 'band'):
                fill, _line, _dash, text, _b, _s, _lw = styles[role]
                self.assertGreaterEqual(_contrast(text, fill), 4.5, (name, role))
            self.assertGreaterEqual(_contrast(styles['note'][3], '#FFFFFF'), 4.5, (name, 'note'))

    def test_set_default_theme_and_cli(self):
        vlib.set_default_theme('mono')
        self.addCleanup(vlib.set_default_theme, 'tech-cn')
        self.assertEqual(Fig(50, 20).theme, 'mono')
        # 显式 theme= 不受默认影响
        self.assertEqual(Fig(50, 20, theme='mono').theme, 'mono')
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, 'out')
            subprocess.run([sys.executable, os.path.join(XB, 'build_vsdx.py'),
                            EXAMPLE, out, '--theme', 'mono'], check=True, capture_output=True)
            with zipfile.ZipFile(os.path.join(out, 'figdemo-1.vsdx')) as z:
                self.assertIn('#000000', z.read('visio/pages/page1.xml').decode())
                self.assertNotIn('#1F3864', z.read('visio/pages/page1.xml').decode())


def make_docx(path, vsdx_bytes):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="emf" ContentType="image/x-emf"/>'
          '<Default Extension="vsdx" ContentType="application/vnd.ms-visio.drawing"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Visio_Drawing.vsdx"/>'
             '</Relationships>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'
           '<w:body><w:p><w:r><w:object w:dxaOrig="5000" w:dyaOrig="3000">'
           '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:300pt;height:100pt" o:ole="">'
           '<v:imagedata r:id="rId2" o:title=""/></v:shape>'
           '<o:OLEObject Type="Embed" ProgID="Visio.Drawing.15" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rId3"/>'
           '</w:object></w:r></w:p></w:body></w:document>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/_rels/document.xml.rels', drels)
        z.writestr('word/media/image1.emf', b'\x01\x00\x00\x00' + b'\x00' * 100)
        z.writestr('word/embeddings/Microsoft_Visio_Drawing.vsdx', vsdx_bytes)


class TestPipeline(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.blank = os.path.join(self.d, 'blank.vsdx')
        vsdxpkg.zipdir(vsdxpkg.BLANK, self.blank)
        with open(self.blank, 'rb') as fh:
            self.blank_bytes = fh.read()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_build_docx_importable(self):
        import build_docx  # 不能在 import 时解析命令行退出
        self.assertTrue(callable(build_docx.main))

    def test_build_vsdx_cli(self):
        out = os.path.join(self.d, 'out')
        subprocess.run([sys.executable, os.path.join(XB, 'build_vsdx.py'), EXAMPLE, out], check=True, capture_output=True)
        self.assertEqual(sorted(os.listdir(out)), ['figdemo-1.png', 'figdemo-1.vsdx'])
        with zipfile.ZipFile(os.path.join(out, 'figdemo-1.vsdx')) as z:
            ET.fromstring(z.read('visio/pages/page1.xml'))
            pages = z.read('visio/pages/pages.xml').decode()
            self.assertIn(f'V="{150/25.4:.9f}"', pages)

    def test_extract_clears_stale_parts(self):
        src = os.path.join(self.d, 'a.docx')
        with zipfile.ZipFile(self.blank) as z, zipfile.ZipFile(os.path.join(self.d, 'extra.vsdx'), 'w') as o:
            for n in z.namelist():
                o.writestr(n, z.read(n))
            o.writestr('extra.txt', 'stale')
        with open(os.path.join(self.d, 'extra.vsdx'), 'rb') as fh:
            make_docx(src, fh.read())
        tpl = os.path.join(self.d, 'tpl')
        import extract_vsdx
        with contextlib.redirect_stdout(io.StringIO()):
            extract_vsdx.main(src, tpl)
            self.assertTrue(os.path.exists(os.path.join(tpl, 'Microsoft_Visio_Drawing', 'extra.txt')))
            make_docx(src, self.blank_bytes)
            extract_vsdx.main(src, tpl)
        self.assertFalse(os.path.exists(os.path.join(tpl, 'Microsoft_Visio_Drawing', 'extra.txt')))

    def test_docx_roundtrip(self):
        src = os.path.join(self.d, 'src.docx')
        make_docx(src, self.blank_bytes)
        tpl, out = os.path.join(self.d, 'tpl'), os.path.join(self.d, 'out.docx')
        import build_docx, extract_vsdx
        with contextlib.redirect_stdout(io.StringIO()):
            extract_vsdx.main(src, tpl)
            build_docx.build(src, EXAMPLE, tpl, out)
        with zipfile.ZipFile(out) as z:
            names = z.namelist()
            doc = z.read('word/document.xml').decode()
            rels = z.read('word/_rels/document.xml.rels').decode()
            with zipfile.ZipFile(io.BytesIO(z.read('word/embeddings/Microsoft_Visio_Drawing.vsdx'))) as v:
                page = v.read('visio/pages/page1.xml').decode()
        self.assertEqual(names[0], '[Content_Types].xml')
        self.assertIn('word/media/figdemo-1.png', names)
        self.assertNotIn('word/media/image1.emf', names)
        self.assertIn('Target="media/figdemo-1.png"', rels)
        self.assertIn(f'width:{150/25.4*72:.2f}pt;height:{70/25.4*72:.2f}pt', doc)
        self.assertIn('aspectratio="t"', doc)
        self.assertIn('参数校验', page)
        env = dict(os.environ, PYTHONIOENCODING='utf-8')
        r = subprocess.run([sys.executable, os.path.join(XB, 'inspect_visio_docx.py'), out], capture_output=True, text=True, encoding='utf-8', env=env, check=True)
        self.assertIn('问题：无', r.stdout)

    def test_build_docx_theme_kwarg(self):
        src = os.path.join(self.d, 'src.docx')
        make_docx(src, self.blank_bytes)
        tpl, out = os.path.join(self.d, 'tpl'), os.path.join(self.d, 'out-mono.docx')
        import build_docx, extract_vsdx, vlib as _vlib
        self.addCleanup(_vlib.set_default_theme, 'tech-cn')
        with contextlib.redirect_stdout(io.StringIO()):
            extract_vsdx.main(src, tpl)
            build_docx.build(src, EXAMPLE, tpl, out, theme='mono')
        with zipfile.ZipFile(out) as z:
            with zipfile.ZipFile(io.BytesIO(z.read('word/embeddings/Microsoft_Visio_Drawing.vsdx'))) as v:
                page = v.read('visio/pages/page1.xml').decode()
        self.assertIn('#000000', page)
        self.assertNotIn('#1F3864', page)

    def test_bad_suffix_message(self):
        src = os.path.join(self.d, 'src.docx')
        make_docx(src, self.blank_bytes)
        spec = os.path.join(self.d, 'specs.py')
        with open(spec, 'w', encoding='utf8') as fh:
            fh.write(f'import sys; sys.path.insert(0, {XB!r})\nfrom vlib import Fig\n'
                     'def f():\n    g = Fig(20, 10); g.box("a", 1, 1, 5, 5); return g\n'
                     'FIGS = [("x", "99", f)]\n')
        import build_docx
        with self.assertRaisesRegex(SystemExit, "'99'"), contextlib.redirect_stdout(io.StringIO()):
            build_docx.build(src, spec, os.path.join(self.d, 'none'), os.path.join(self.d, 'o.docx'))

    def test_multipage_template_rejected(self):
        tpl = os.path.join(self.d, 'mp')
        shutil.copytree(vsdxpkg.BLANK, tpl)
        pp = os.path.join(tpl, 'visio', 'pages', 'pages.xml')
        with open(pp, encoding='utf8') as fh:
            s = fh.read()
        page = re.search(r'<Page\b.*?</Page>', s, flags=re.S).group(0)
        with open(pp, 'w', encoding='utf8') as fh:
            fh.write(s.replace(page, page + page))
        with self.assertRaises(ValueError):
            vsdxpkg.build_vsdx(example_fig(), tpl, os.path.join(self.d, 'x.vsdx'))


if __name__ == '__main__':
    unittest.main()
