#!/usr/bin/env python3
"""从 docx 中解出全部嵌入的 Visio 图，作为该文档的模板目录（保留每张图自己的样式表等部件）。

用法：python3 extract_vsdx.py <原稿.docx> <输出目录>
输出：<输出目录>/Microsoft_Visio_DrawingN/（解包目录，供 build_docx.py 用），
      并打印“正文顺序 → 嵌入文件”对照表，用来填写 figspecs.py 里的 FIGS。
只读原稿。依赖：Python 3 标准库。
"""
import os, re, shutil, sys, zipfile


def main(docx, out):
    os.makedirs(out, exist_ok=True)
    with zipfile.ZipFile(docx) as z:
        rels = z.read('word/_rels/document.xml.rels').decode('utf8')
        doc = z.read('word/document.xml').decode('utf8')
        rid = {re.search(r'Id="([^"]+)"', r).group(1): re.search(r'Target="([^"]+)"', r).group(1)
               for r in re.findall(r'<Relationship [^>]*?/>', rels)}
        n = 0
        for i, m in enumerate(re.finditer(r'<w:object [^>]*>.*?</w:object>', doc, flags=re.S), 1):
            ids = re.findall(r'r:id="(rId\d+)"', m.group(0))
            emb = [rid[x] for x in ids if x in rid and rid[x].lower().endswith('.vsdx')]
            if not emb:
                continue
            name = os.path.basename(emb[0])
            dst = os.path.join(out, name[:-5])
            shutil.rmtree(dst, ignore_errors=True)   # 先清空，避免上次解出的旧部件残留进新 vsdx
            with zipfile.ZipFile(z.open('word/' + emb[0])) as v:
                v.extractall(dst)
            n += 1
            suf = name[len('Microsoft_Visio_Drawing'):-5]
            print(f'#{n:<3} suffix={suf!r:<6} {name}')
    print(f'共 {n} 个 vsdx -> {out}')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(*sys.argv[1:])
