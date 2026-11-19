"""Tiny diagram DSL -> Visio page XML + PNG preview. Units mm, origin top-left, y down."""
import math
import os
from PIL import Image, ImageDraw, ImageFont
from xml.sax.saxutils import escape

def _find_fonts():
    """PNG 预览用的中文字体（Visio XML 里始终写微软雅黑）。
    查找顺序：环境变量 VISIO_KIT_FONT_R / VISIO_KIT_FONT_B → Windows 微软雅黑 → 常见 Noto/思源 CJK。"""
    env_r, env_b = os.environ.get('VISIO_KIT_FONT_R'), os.environ.get('VISIO_KIT_FONT_B')
    if env_r:
        return env_r, env_b or env_r
    cands = [
        ('C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/msyhbd.ttc'),
        ('/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/PingFang.ttc'),
        ('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'),
        ('/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc', '/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc'),
        ('/usr/share/fonts/google-noto-cjk/NotoSansCJK-Regular.ttc', '/usr/share/fonts/google-noto-cjk/NotoSansCJK-Bold.ttc'),
    ]
    for r, b in cands:
        if os.path.exists(r):
            return r, b if os.path.exists(b) else r
    raise RuntimeError('找不到中文字体：请设置环境变量 VISIO_KIT_FONT_R（及可选 VISIO_KIT_FONT_B）指向 .ttf/.otf/.ttc')


FONT_R = FONT_B = None


def _fonts():
    global FONT_R, FONT_B
    if FONT_R is None:
        FONT_R, FONT_B = _find_fonts()
    return FONT_R, FONT_B
LATIN = 'Microsoft YaHei'
ASIAN = '微软雅黑'

# palette
NAVY = '#1F3864'
BLUE = '#2F5597'
LINE = '#2E5C8A'
GRAY = '#7F7F7F'
TXT = '#1A1A1A'

STYLES = {
    #        fill      line     dash  text     bold  size  lw
    'node':  ('#FFFFFF', LINE,  0, TXT,      0, 9.0, 0.75),
    'key':   ('#DCE8F5', LINE,  0, NAVY,     1, 9.0, 0.75),
    'dark':  (NAVY,     NAVY,   0, '#FFFFFF', 1, 9.0, 0.75),
    'ext':   ('#F2F2F2', GRAY,  1, '#404040', 0, 9.0, 0.75),
    'ok':    ('#E2F0D9', '#548235', 0, '#385723', 0, 9.0, 0.75),
    'warn':  ('#FBE5D6', '#C55A11', 0, '#843C0C', 0, 9.0, 0.75),
    'soft':  ('#F3F6FA', '#A6B9D0', 0, '#1F3864', 0, 8.5, 0.5),
    'text':  (None,      None,   0, TXT,      0, 9.0, 0),
    'title': (None,      None,   0, NAVY,     1, 10.0, 0),
    'note':  (None,      None,   0, '#595959', 0, 8.0, 0),
    'label': ('#FFFFFF', None,   0, '#404040', 0, 8.0, 0),
    'band':  (BLUE,      BLUE,   0, '#FFFFFF', 1, 9.0, 0.75),
}

class Fig:
    def __init__(self, W, H):
        self.W, self.H = W, H
        self.items = []  # draw order
        self.boxes = {}

    # ---- primitives
    def box(self, id, x, y, w, h, text='', style='node', align='c', r=0, size=None, bold=None, valign='m'):
        b = dict(kind='box', id=id, x=x, y=y, w=w, h=h, text=text, style=style, align=align, r=r,
                 size=size, bold=bold, valign=valign)
        self.items.append(b)
        if id:
            self.boxes[id] = b
        return b

    def text(self, x, y, w, h, text, style='text', align='c', size=None, bold=None, valign='m'):
        return self.box(None, x, y, w, h, text, style, align, 0, size, bold, valign)

    def container(self, id, x, y, w, h, title, dashed=False, band=True, tw=None):
        c = dict(kind='container', id=id, x=x, y=y, w=w, h=h, title=title, dashed=dashed, band=band)
        self.items.append(c)
        self.boxes[id] = dict(x=x, y=y, w=w, h=h)
        if title and band:
            bw = tw or max(30, len(title) * 3.6 + 8)
            self.box(None, x, y, bw, 6.5, title, 'band', 'c', 0)
        elif title:
            self.box(None, x + 2, y + 1, w - 4, 6, title, 'title', 'l', 0)
        return c

    def anchor(self, id, side, f=0.5):
        b = self.boxes[id]
        x, y, w, h = b['x'], b['y'], b['w'], b['h']
        return {'t': (x + w * f, y), 'b': (x + w * f, y + h), 'l': (x, y + h * f), 'r': (x + w, y + h * f)}[side]

    def line(self, pts, dashed=False, arrow='end', color=None, label=None, lpos=None, lw=1.0, src=None, dst=None):
        # label is the connector's own text (Visio connector Text), not a separate text shape
        if label and lpos is None:
            best = max(range(len(pts) - 1), key=lambda i: math.dist(pts[i], pts[i + 1]))
            (x1, y1), (x2, y2) = pts[best], pts[best + 1]
            lpos = ((x1 + x2) / 2, (y1 + y2) / 2)
        l = dict(kind='line', pts=pts, dashed=dashed, arrow=arrow, color=color or (GRAY if dashed else LINE), lw=lw,
                 src=src, dst=dst, label=label, lpos=lpos)
        self.items.append(l)
        return l

    def conn(self, a, sa, b, sb, fa=0.5, fb=0.5, via=None, mid=None, **kw):
        p0 = self.anchor(a, sa, fa)
        p1 = self.anchor(b, sb, fb)
        if via is not None:
            pts = [p0] + via + [p1]
        elif sa in 'tb' and sb in 'tb':
            if abs(p0[0] - p1[0]) < 0.05:
                pts = [p0, p1]
            else:
                my = mid if mid is not None else (p0[1] + p1[1]) / 2
                pts = [p0, (p0[0], my), (p1[0], my), p1]
        elif sa in 'lr' and sb in 'lr':
            if abs(p0[1] - p1[1]) < 0.05:
                pts = [p0, p1]
            else:
                mx = mid if mid is not None else (p0[0] + p1[0]) / 2
                pts = [p0, (mx, p0[1]), (mx, p1[1]), p1]
        elif sa in 'tb':
            pts = [p0, (p0[0], p1[1]), p1]
        else:
            pts = [p0, (p1[0], p0[1]), p1]
        return self.line(pts, src=a, dst=b, **kw)

    def legend(self, x, y, entries):
        # entries: list of (label, dashed, color)
        cx = x
        for lab, dashed, col in entries:
            self.line([(cx, y), (cx + 10, y)], dashed=dashed, color=col)
            self.text(cx + 11.5, y - 2.5, textwidth(lab, 8) + 2, 5, lab, 'note', 'l')
            cx += 11.5 + textwidth(lab, 8) + 8

    # ---- outputs
    def render_png(self, path, S=9.0):
        W, H = int(self.W * S), int(self.H * S)
        img = Image.new('RGB', (W, H), 'white')
        d = ImageDraw.Draw(img)
        P = lambda p: (p[0] * S, p[1] * S)
        for it in self.items:
            if it['kind'] == 'container':
                x0, y0, x1, y1 = it['x'] * S, it['y'] * S, (it['x'] + it['w']) * S, (it['y'] + it['h']) * S
                if it['dashed']:
                    d.rectangle([x0, y0, x1, y1], fill='#FFFFFF')
                    dashed_poly(d, [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)], '#8EA9C1', max(1, int(0.75 * 0.3528 * S)), S)
                else:
                    d.rectangle([x0, y0, x1, y1], fill='#F7F9FC', outline='#8EA9C1', width=max(1, int(0.75 * 0.3528 * S)))
            elif it['kind'] == 'box':
                fill, line, dash, tc, bold, size, lw = STYLES[it['style']]
                x0, y0, x1, y1 = it['x'] * S, it['y'] * S, (it['x'] + it['w']) * S, (it['y'] + it['h']) * S
                rr = it['r'] * S
                wpx = max(1, round(lw * 0.3528 * S))
                if fill:
                    d.rounded_rectangle([x0, y0, x1, y1], radius=rr, fill=fill)
                if line:
                    if dash:
                        dashed_poly(d, [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)], line, wpx, S)
                    else:
                        d.rounded_rectangle([x0, y0, x1, y1], radius=rr, outline=line, width=wpx)
                if it['text']:
                    sz = it['size'] or size
                    bd = it['bold'] if it['bold'] is not None else bold
                    fnt = ImageFont.truetype(_fonts()[1] if bd else _fonts()[0], int(round(sz * 0.3528 * S)))
                    lines = it['text'].split('\n')
                    lh = sz * 0.3528 * 1.25 * S
                    tot = lh * len(lines)
                    if it['valign'] == 't':
                        ty = y0 + 1.2 * S
                    else:
                        ty = (y0 + y1) / 2 - tot / 2
                    for i, ln in enumerate(lines):
                        yy = ty + lh * (i + 0.5)
                        if it['align'] == 'l':
                            d.text((x0 + 1.5 * S, yy), ln, font=fnt, fill=tc, anchor='lm')
                        else:
                            d.text(((x0 + x1) / 2, yy), ln, font=fnt, fill=tc, anchor='mm')
            elif it['kind'] == 'line':
                pts = [P(p) for p in it['pts']]
                wpx = max(1, round(it['lw'] * 0.3528 * S * 1.0))
                if it['dashed']:
                    dashed_poly(d, pts, it['color'], wpx, S)
                else:
                    d.line(pts, fill=it['color'], width=wpx, joint='curve')
                ah = 2.2 * S
                if it.get('label'):
                    fnt = ImageFont.truetype(_fonts()[0], int(round(8.0 * 0.3528 * S)))
                    lx, ly = P(it['lpos'])
                    tw = (textwidth(it['label'], 8.0) + 2.5) * S
                    d.rectangle([lx - tw / 2, ly - 2.4 * S, lx + tw / 2, ly + 2.4 * S], fill='#FFFFFF')
                    d.text((lx, ly), it['label'], font=fnt, fill='#404040', anchor='mm')
                if it['arrow'] in ('end', 'both'):
                    arrowhead(d, pts[-2], pts[-1], ah, it['color'])
                if it['arrow'] in ('start', 'both'):
                    arrowhead(d, pts[1], pts[0], ah, it['color'])
        img.save(path, dpi=(int(S * 25.4), int(S * 25.4)))

    # ---------- glue helpers ----------
    def _glue_targets(self):
        """Shapes a connector end may attach to: bordered boxes and containers (not plain text)."""
        tg = []
        for k, it in enumerate(self.items):
            if it['kind'] == 'container' or (it['kind'] == 'box' and STYLES[it['style']][1]):
                tg.append(k)
        return tg

    def _on_border(self, it, p, TOL=0.3):
        x, y, w, h = it['x'], it['y'], it['w'], it['h']
        on_v = (abs(p[0] - x) < TOL or abs(p[0] - x - w) < TOL) and y - TOL <= p[1] <= y + h + TOL
        on_h = (abs(p[1] - y) < TOL or abs(p[1] - y - h) < TOL) and x - TOL <= p[0] <= x + w + TOL
        return on_v or on_h

    def _find_target(self, p, named):
        if named is not None:
            for k, it in enumerate(self.items):
                if it['kind'] in ('box', 'container') and it.get('id') == named and self._on_border(it, p):
                    return k
        best = None
        for k in self._glue_targets():
            it = self.items[k]
            if self._on_border(it, p):
                if best is None or it['w'] * it['h'] < self.items[best]['w'] * self.items[best]['h']:
                    best = k
        return best

    def glue_report(self):
        n_end = n_glued = 0
        for it in self.items:
            if it['kind'] != 'line':
                continue
            for p, nm in ((it['pts'][0], it['src']), (it['pts'][-1], it['dst'])):
                n_end += 1
                if self._find_target(p, nm) is not None:
                    n_glued += 1
        return n_glued, n_end

    def visio_xml(self):
        H = self.H
        I = lambda mm: mm / 25.4
        sid = [0]
        out = []
        connects = []
        shape_of = {}          # item index -> Visio sheet id
        cpts = {}              # item index -> list of (fx, fy_from_top)

        def cp_index(k, p):
            it = self.items[k]
            fx = (p[0] - it['x']) / it['w']
            fy = (p[1] - it['y']) / it['h']
            fx, fy = min(max(fx, 0), 1), min(max(fy, 0), 1)
            lst = cpts.setdefault(k, [(0.5, 0), (0.5, 1), (0, 0.5), (1, 0.5)])
            for j, (a, b) in enumerate(lst):
                if abs(a - fx) < 1e-4 and abs(b - fy) < 1e-4:
                    return j
            lst.append((fx, fy))
            return len(lst) - 1

        # pass 1: resolve every connector end to (target item, connection-point row)
        ends = {}
        for k, it in enumerate(self.items):
            if it['kind'] != 'line':
                continue
            r = []
            for p, nm in ((it['pts'][0], it['src']), (it['pts'][-1], it['dst'])):
                t = self._find_target(p, nm)
                r.append(None if t is None else (t, cp_index(t, p)))
            ends[k] = r
        for k in self._glue_targets():
            cpts.setdefault(k, [(0.5, 0), (0.5, 1), (0, 0.5), (1, 0.5)])

        def nid():
            sid[0] += 1
            return sid[0]

        def char(color, size, bold):
            return (f'<Section N="Character"><Row IX="0"><Cell N="Font" V="{LATIN}"/><Cell N="Color" V="{color}"/>'
                    f'<Cell N="Style" V="{1 if bold else 0}"/><Cell N="Size" V="{size/72:.6f}" U="PT"/>'
                    f'<Cell N="AsianFont" V="{ASIAN}"/><Cell N="LangID" V="zh-CN"/></Row></Section>')

        def rect_geom(w, h, nofill, noline):
            return (f'<Section N="Geometry" IX="0"><Cell N="NoFill" V="{nofill}"/><Cell N="NoLine" V="{noline}"/>'
                    f'<Cell N="NoShow" V="0"/><Cell N="NoSnap" V="0"/>'
                    f'<Row T="MoveTo" IX="1"><Cell N="X" V="0" F="Width*0"/><Cell N="Y" V="0" F="Height*0"/></Row>'
                    f'<Row T="LineTo" IX="2"><Cell N="X" V="{w}" F="Width*1"/><Cell N="Y" V="0" F="Height*0"/></Row>'
                    f'<Row T="LineTo" IX="3"><Cell N="X" V="{w}" F="Width*1"/><Cell N="Y" V="{h}" F="Height*1"/></Row>'
                    f'<Row T="LineTo" IX="4"><Cell N="X" V="0" F="Width*0"/><Cell N="Y" V="{h}" F="Height*1"/></Row>'
                    f'<Row T="LineTo" IX="5"><Cell N="X" V="0" F="Geometry1.X1"/><Cell N="Y" V="0" F="Geometry1.Y1"/></Row></Section>')

        def conn_section(wi, hi, lst):
            rows = ''.join(
                f'<Row IX="{j}"><Cell N="X" V="{wi*fx:.6f}" F="Width*{fx:.6f}"/><Cell N="Y" V="{hi*(1-fy):.6f}" F="Height*{1-fy:.6f}"/>'
                f'<Cell N="DirX" V="0"/><Cell N="DirY" V="0"/><Cell N="Type" V="0"/><Cell N="AutoGen" V="0"/></Row>'
                for j, (fx, fy) in enumerate(lst))
            return f'<Section N="Connection">{rows}</Section>'

        def rect_shape(k, x, y, w, h, fill, line, dash, lw, text='', tc=TXT, size=9, bold=0, align='c', r=0, valign='m'):
            i = nid()
            shape_of[k] = i
            wi, hi = I(w), I(h)
            cx, cy = I(x + w / 2), I(H - y - h / 2)
            cells = [f'<Cell N="PinX" V="{cx}"/>', f'<Cell N="PinY" V="{cy}"/>', f'<Cell N="Width" V="{wi}"/>',
                     f'<Cell N="Height" V="{hi}"/>', f'<Cell N="LocPinX" V="{wi/2}" F="Width*0.5"/>',
                     f'<Cell N="LocPinY" V="{hi/2}" F="Height*0.5"/>', '<Cell N="Angle" V="0"/>',
                     f'<Cell N="FillForegnd" V="{fill or "#FFFFFF"}"/>', f'<Cell N="FillPattern" V="{1 if fill else 0}"/>',
                     f'<Cell N="LineColor" V="{line or "#000000"}"/>',
                     f'<Cell N="LinePattern" V="{0 if not line else (2 if dash else 1)}"/>',
                     f'<Cell N="LineWeight" V="{(lw or 0.75)/72:.6f}" U="PT"/>',
                     f'<Cell N="Rounding" V="{I(r):.6f}" U="MM"/>', '<Cell N="ShdwPattern" V="0"/>',
                     '<Cell N="LeftMargin" V="0.0394" U="MM"/>', '<Cell N="RightMargin" V="0.0394" U="MM"/>',
                     '<Cell N="TopMargin" V="0.0394" U="MM"/>', '<Cell N="BottomMargin" V="0.0394" U="MM"/>',
                     f'<Cell N="VerticalAlign" V="{0 if valign=="t" else 1}"/>']
            body = ''.join(cells)
            body += char(tc, size, bold)
            body += f'<Section N="Paragraph"><Row IX="0"><Cell N="HorzAlign" V="{0 if align=="l" else 1}"/><Cell N="SpLine" V="-1.15" U="PER"/></Row></Section>'
            if k in cpts:
                body += conn_section(wi, hi, cpts[k])
            body += rect_geom(wi, hi, 0 if fill else 1, 0 if line else 1)
            if text:
                body += '<Text><cp IX="0"/><pp IX="0"/>' + escape(text) + '</Text>'
            out.append(f'<Shape ID="{i}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">{body}</Shape>')
            return i

        # shapes first, so connector formulas can reference their sheet ids
        for k, it in enumerate(self.items):
            if it['kind'] == 'container':
                rect_shape(k, it['x'], it['y'], it['w'], it['h'], None if it['dashed'] else '#F7F9FC', '#8EA9C1',
                           it['dashed'], 0.75)
            elif it['kind'] == 'box':
                fill, line, dash, tc, bold, size, lw = STYLES[it['style']]
                bd = it['bold'] if it['bold'] is not None else bold
                rect_shape(k, it['x'], it['y'], it['w'], it['h'], fill, line, dash, lw, it['text'], tc,
                           it['size'] or size, bd, it['align'], it['r'], it['valign'])

        for k, it in enumerate(self.items):
            if it['kind'] != 'line':
                continue
            i = nid()
            pts = [(I(x), I(H - y)) for x, y in it['pts']]
            (bx, by), (ex, ey) = pts[0], pts[-1]
            W_, H_ = ex - bx, ey - by
            rows = []
            for j, (px, py) in enumerate(pts):
                lx, ly = px - bx, py - by
                fxs = f' F="Width*{lx / W_:.6f}"' if abs(W_) > 1e-6 else ''
                fys = f' F="Height*{ly / H_:.6f}"' if abs(H_) > 1e-6 else ''
                rows.append(f'<Row T="{"MoveTo" if j == 0 else "LineTo"}" IX="{j+1}">'
                            f'<Cell N="X" V="{lx:.6f}"{fxs}/><Cell N="Y" V="{ly:.6f}"{fys}/></Row>')
            # ends: glue to a connection point of the shape at that end
            endcells = []
            for (cx_, cy_, nmx, nmy, part), e in zip(((bx, by, 'BeginX', 'BeginY', 9), (ex, ey, 'EndX', 'EndY', 12)), ends[k]):
                if e is None:
                    endcells.append(f'<Cell N="{nmx}" V="{cx_}"/><Cell N="{nmy}" V="{cy_}"/>')
                else:
                    t, j = e
                    ref = f'PAR(PNT(Sheet.{shape_of[t]}!Connections.X{j+1},Sheet.{shape_of[t]}!Connections.Y{j+1}))'
                    endcells.append(f'<Cell N="{nmx}" V="{cx_}" F="{ref}"/><Cell N="{nmy}" V="{cy_}" F="{ref}"/>')
                    connects.append(f'<Connect FromSheet="{i}" FromCell="{nmx}" FromPart="{part}" '
                                    f'ToSheet="{shape_of[t]}" ToCell="Connections.X{j+1}" ToPart="{100 + j}"/>')
            arr = it['arrow']
            body = (f'<Cell N="PinX" V="{(bx+ex)/2}" F="(BeginX+EndX)/2"/><Cell N="PinY" V="{(by+ey)/2}" F="(BeginY+EndY)/2"/>'
                    f'<Cell N="Width" V="{W_}" F="EndX-BeginX"/><Cell N="Height" V="{H_}" F="EndY-BeginY"/>'
                    f'<Cell N="LocPinX" V="{W_/2}" F="Width*0.5"/><Cell N="LocPinY" V="{H_/2}" F="Height*0.5"/>'
                    f'<Cell N="Angle" V="0"/>' + ''.join(endcells) +
                    f'<Cell N="LineColor" V="{it["color"]}"/><Cell N="LinePattern" V="{2 if it["dashed"] else 1}"/>'
                    f'<Cell N="LineWeight" V="{it["lw"]/72:.6f}" U="PT"/><Cell N="Rounding" V="0"/>'
                    f'<Cell N="BeginArrow" V="{13 if arr in ("start","both") else 0}"/>'
                    f'<Cell N="EndArrow" V="{13 if arr in ("end","both") else 0}"/>'
                    f'<Cell N="BeginArrowSize" V="1"/><Cell N="EndArrowSize" V="1"/><Cell N="FillPattern" V="0"/>'
                    f'<Cell N="ObjType" V="2"/><Cell N="ShapeRouteStyle" V="1"/><Cell N="ConLineRouteExt" V="1"/>'
                    f'<Cell N="ConFixedCode" V="0"/><Cell N="ShapePlowCode" V="0"/><Cell N="GlueType" V="0"/>'
                    f'<Cell N="NoAlignBox" V="1"/><Cell N="DynFeedback" V="0"/>')
            if it.get('label'):
                lab = it['label']
                lx, ly = it['lpos']
                tw, th = I(textwidth(lab, 8.0) + 2.5), I(4.8)
                tx, ty = I(lx) - bx, I(H - ly) - by
                txs = f' F="Width*{tx / W_:.6f}"' if abs(W_) > 1e-6 else ''
                tys = f' F="Height*{ty / H_:.6f}"' if abs(H_) > 1e-6 else ''
                body += (f'<Cell N="TxtPinX" V="{tx:.6f}"{txs}/><Cell N="TxtPinY" V="{ty:.6f}"{tys}/>'
                         f'<Cell N="TxtWidth" V="{tw:.6f}"/><Cell N="TxtHeight" V="{th:.6f}"/>'
                         f'<Cell N="TxtLocPinX" V="{tw/2:.6f}" F="TxtWidth*0.5"/><Cell N="TxtLocPinY" V="{th/2:.6f}" F="TxtHeight*0.5"/>'
                         f'<Cell N="TxtAngle" V="0"/><Cell N="TextBkgnd" V="#FFFFFF" F="RGB(255,255,255)"/>'
                         '<Cell N="TextBkgndTrans" V="0"/><Cell N="VerticalAlign" V="1"/>'
                         '<Cell N="LeftMargin" V="0"/><Cell N="RightMargin" V="0"/><Cell N="TopMargin" V="0"/><Cell N="BottomMargin" V="0"/>')
                body += char('#404040', 8.0, 0)
                body += '<Section N="Paragraph"><Row IX="0"><Cell N="HorzAlign" V="1"/></Row></Section>'
            body += (f'<Section N="Geometry" IX="0"><Cell N="NoFill" V="1"/><Cell N="NoLine" V="0"/><Cell N="NoShow" V="0"/>'
                     + ''.join(rows) + '</Section>')
            if it.get('label'):
                body += '<Text><cp IX="0"/><pp IX="0"/>' + escape(it['label']) + '</Text>'
            out.append(f'<Shape ID="{i}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">{body}</Shape>')
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xml:space="preserve">'
               '<Shapes>' + ''.join(out) + '</Shapes>')
        if connects:
            xml += '<Connects>' + ''.join(connects) + '</Connects>'
        xml += '</PageContents>'
        return xml


def textwidth(s, pt):
    w = 0
    for ch in s:
        w += 1.0 if ord(ch) > 0x2E80 else 0.56
    return w * pt * 0.3528


def arrowhead(d, p, q, L, col):
    a = math.atan2(q[1] - p[1], q[0] - p[0])
    d.polygon([q, (q[0] - L * math.cos(a - 0.38), q[1] - L * math.sin(a - 0.38)),
               (q[0] - L * math.cos(a + 0.38), q[1] - L * math.sin(a + 0.38))], fill=col)


def dashed_poly(d, pts, col, w, S):
    dash, gap = 2.0 * S, 1.2 * S
    for i in range(len(pts) - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        L = math.hypot(x2 - x1, y2 - y1)
        if L == 0:
            continue
        t = 0
        while t < L:
            t2 = min(L, t + dash)
            d.line([x1 + (x2 - x1) * t / L, y1 + (y2 - y1) * t / L, x1 + (x2 - x1) * t2 / L, y1 + (y2 - y1) * t2 / L], fill=col, width=w)
            t = t2 + gap
