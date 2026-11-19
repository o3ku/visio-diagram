# -*- coding: utf-8 -*-
"""原地替换 docx 里**已有的** Visio OLE 图（不动容器、不重排文档）。

甲方文档里常见的形态：图是 `word/embeddings/*.vsdx`，Word 显示的却是缓存的
`word/media/*.emf` 预览。只换一处的结果是「显示的图」和「双击进去的图」不一致 ——
必须两处一起换；显示框宽高也要按新 EMF 的宽高比重写，否则 Word 按旧框拉伸变扁。

用法：
    python replace_ole.py 方案.docx 图N2.emf 图N2.vsdx [--width-pt 415]

不给 --width-pt 时按正文栏宽（末节 `w:pgMar` 实算）留 0.5pt 余量。
"""
import argparse
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_embed                                        # noqa: E402

OBJ_RE = re.compile(r"<w:object\b[^>]*>.*?</w:object>", re.S)
STYLE_RE = re.compile(r'(style="width:)([\d.]+)(pt;height:)([\d.]+)(pt)')


def emf_aspect(path):
    """从 EMF 头 rclFrame（单位 0.01 mm）取宽高比 —— 唯一可信的图幅比例。"""
    data = open(path, "rb").read()
    i = data.find(b" EMF")
    if i < 0:
        raise SystemExit("不是 EMF：%s" % path)
    l, t, r, b = (int.from_bytes(data[i + 24 + 4 * k:i + 28 + 4 * k], "little")
                  for k in range(4))
    if r - l <= 0 or b - t <= 0:
        raise SystemExit("EMF 图幅为空（rclFrame）：%s" % path)
    return (r - l) / float(b - t)


def find_slots(names):
    emf = [n for n in names if n.startswith("word/media/") and n.lower().endswith(".emf")]
    ole = [n for n in names if n.startswith("word/embeddings/") and n.lower().endswith(".vsdx")]
    if len(emf) != 1 or len(ole) != 1:
        raise SystemExit("文档里 EMF %d 个 / 嵌入 vsdx %d 个，本脚本只处理「一图一对象」；"
                         "多个时请手工指定槽位" % (len(emf), len(ole)))
    return emf[0], ole[0]


def main():
    ap = argparse.ArgumentParser(description="原地替换 docx 内嵌 Visio 图")
    ap.add_argument("docx")
    ap.add_argument("emf")
    ap.add_argument("vsdx")
    ap.add_argument("--width-pt", type=float, default=None,
                    help="显示框宽度（pt）；默认取正文栏宽 −0.5pt")
    a = ap.parse_args()

    aspect = emf_aspect(a.emf)
    if a.width_pt:
        width = a.width_pt
    else:
        col_mm, _ = docx_embed.content_column_mm(a.docx)     # 返回 (mm, twips)
        width = round(col_mm * 72 / 25.4 - 0.5, 1)
    height = round(width / aspect, 2)

    with zipfile.ZipFile(a.docx) as z:
        names = z.namelist()
        emf_slot, ole_slot = find_slots(names)

    tmp = a.docx + ".tmp"
    patched = 0
    with zipfile.ZipFile(a.docx) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename == emf_slot:
                dst.writestr(item, open(a.emf, "rb").read())
            elif item.filename == ole_slot:
                dst.writestr(item, open(a.vsdx, "rb").read())
            elif item.filename == "word/document.xml":
                xml = src.read(item.filename).decode("utf-8")

                def fix(m):
                    nonlocal patched
                    patched += 1
                    return STYLE_RE.sub(lambda s: "%s%g%s%g%s" % (
                        s.group(1), width, s.group(3), height, s.group(5)), m.group(0), count=1)

                xml = OBJ_RE.sub(fix, xml)
                dst.writestr(item, xml.encode("utf-8"))
            else:
                dst.writestr(item, src.read(item.filename))
    if not patched:
        os.remove(tmp)
        raise SystemExit("document.xml 里没找到 <w:object> 显示框，未改动文档")
    shutil.move(tmp, a.docx)
    print("已替换：%s\n  %s ← %s\n  %s ← %s\n  显示框 %.1f x %.1f pt（比 %.4f）"
          % (os.path.basename(a.docx), emf_slot, os.path.basename(a.emf),
             ole_slot, os.path.basename(a.vsdx), width, height, aspect))


if __name__ == "__main__":
    main()
