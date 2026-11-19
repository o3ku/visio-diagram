"""XML 后端示例：一张带判定分支的流程图（不需要 Visio，任何平台可运行）。

生成独立 vsdx + PNG：
    python scripts/xml_backend/build_vsdx.py assets/example_xml_flowchart.py out/
写回 docx（替换已有嵌入图）：见 references/cross-platform.md
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts', 'xml_backend'))
from vlib import Fig  # noqa: E402


def flow():
    f = Fig(150, 70)                                  # 画布 mm，左上原点，Y 向下
    f.box('start', 6, 8, 28, 10, '开始', r=2)          # 只有开始/结束用圆角
    f.box('read', 46, 8, 36, 10, '读取配置')
    f.box('chk', 94, 8, 36, 10, '参数校验', 'key')
    f.box('run', 94, 50, 36, 10, '执行任务', 'ok')
    f.box('err', 46, 50, 36, 10, '报错退出', 'warn')
    f.conn('start', 'r', 'read', 'l')
    f.conn('read', 'r', 'chk', 'l')
    f.conn('chk', 'b', 'run', 't', label='通过')          # 连线文字写在连线自身
    f.conn('chk', 'b', 'err', 't', fa=0.3, mid=30, dashed=True, label='不通过')
    return f


# (图号, 嵌入后缀, 生成函数)；后缀对应 docx 里的 Microsoft_Visio_Drawing{后缀}.vsdx
FIGS = [('demo-1', '', flow)]
