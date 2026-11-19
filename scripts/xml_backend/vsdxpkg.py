"""vsdx 打包工具：解包 / 用模板目录 + 新 page1.xml 生成 vsdx。依赖：Python 3 标准库。"""
import os, re, shutil, tempfile, zipfile

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BLANK = os.path.join(REPO, 'assets', 'templates', 'blank_vsdx')


def zipdir(src, dst):
    files = []
    for root, _, fs in os.walk(src):
        for f in fs:
            files.append(os.path.relpath(os.path.join(root, f), src).replace(os.sep, '/'))
    files.sort(key=lambda p: (p != '[Content_Types].xml', p))
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(os.path.join(src, f), f)


def unzip(src, dst):
    shutil.rmtree(dst, ignore_errors=True)
    with zipfile.ZipFile(src) as z:
        z.extractall(dst)


def build_vsdx(fig, template_dir, out_path, work_dir=None):
    """把 vlib.Fig 写进模板目录的副本（替换 page1.xml，按图尺寸改页面大小），打包为 out_path。
    只支持单页模板；多页模板会报错，避免其余页原样残留。work_dir 缺省用临时目录，用完即删。"""
    own = work_dir is None
    if own:
        work_dir = tempfile.mkdtemp(prefix='visio_xml_vsdx_')
    try:
        shutil.rmtree(work_dir, ignore_errors=True)
        shutil.copytree(template_dir, work_dir)
        pp = os.path.join(work_dir, 'visio', 'pages', 'pages.xml')
        with open(pp, encoding='utf8') as fh:
            pg = fh.read()
        n_pages = len(re.findall(r'<Page\b', pg))
        if n_pages != 1:
            raise ValueError(f'模板 {template_dir} 有 {n_pages} 页；XML 后端只重写 page1，请换单页模板')
        with open(os.path.join(work_dir, 'visio', 'pages', 'page1.xml'), 'w', encoding='utf8') as fh:
            fh.write(fig.visio_xml())
        pg = re.sub(r'(<Cell N="PageWidth" V=")[^"]*', lambda m: m.group(1) + f'{fig.W/25.4:.9f}', pg)
        pg = re.sub(r'(<Cell N="PageHeight" V=")[^"]*', lambda m: m.group(1) + f'{fig.H/25.4:.9f}', pg)
        with open(pp, 'w', encoding='utf8') as fh:
            fh.write(pg)
        zipdir(work_dir, out_path)
    finally:
        if own:
            shutil.rmtree(work_dir, ignore_errors=True)
