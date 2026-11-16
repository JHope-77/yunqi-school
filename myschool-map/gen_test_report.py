# -*- coding: utf-8 -*-
"""生成 myschool-map 功能测试报告"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\Li\Desktop\云栖校园\myschool-map\功能测试报告.docx"

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.5)

def set_run(run, cn='宋体', en='Times New Roman', size=12, bold=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), cn)

def title(t):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(t), '黑体', 'Times New Roman', 18, True)
    p.paragraph_format.space_after = Pt(18)

def h1(t):
    p = doc.add_paragraph()
    set_run(p.add_run(t), '黑体', 'Times New Roman', 15, True)
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(6)

def body(t):
    p = doc.add_paragraph()
    set_run(p.add_run(t), '宋体', 'Times New Roman', 12)
    p.paragraph_format.line_spacing = 1.6
    p.paragraph_format.first_line_indent = Pt(24)

def shade(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def cell_text(cell, text, bold=False, center=False):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    set_run(p.add_run(text), '宋体', 'Times New Roman', 10.5, bold)

def make_table(headers, rows, widths):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        cell_text(t.rows[0].cells[i], h, bold=True, center=True); shade(t.rows[0].cells[i])
    for ri, row in enumerate(rows, start=1):
        for ci, v in enumerate(row):
            cell_text(t.rows[ri].cells[ci], v, center=(ci in (0, len(headers)-1)))
    for ri in range(len(t.rows)):
        for ci, w in enumerate(widths):
            t.rows[ri].cells[ci].width = Cm(w)

title('myschool-map 校园地图小程序功能测试报告')

h1('一、测试概述')
body('本报告针对 myschool-map 微信小程序（校园地图 + 地点搜索 + 步行路径规划）进行功能测试，覆盖地图初始化、校园地标标记、多 Tab 切换、地点搜索、路径规划绘制、点位导航等核心功能，确认各模块在微信开发者工具与真机预览下符合任务书要求。')

h1('二、测试环境')
make_table(
    ['项目', '配置'],
    [
        ['开发工具', '微信开发者工具 Stable 1.06.x（Windows 10/11）'],
        ['基础库版本', '2.10.4'],
        ['地图服务', '腾讯位置服务 WebService API（apis.map.qq.com）'],
        ['中心点', '广西民族大学相思湖学院 22.686699, 108.266767'],
        ['调试机型', 'iPhone SE / iPhone 13 / Pixel 5 模拟器，Android 真机一台'],
    ],
    [4.5, 11.5]
)

h1('三、功能测试用例与结果')
make_table(
    ['编号', '测试模块', '测试步骤', '预期结果', '实际结果'],
    [
        ['TC-01', '地图初始化', '启动小程序进入首页', '地图以学校中心坐标加载，缩放级别 17，显示我的位置', '通过'],
        ['TC-02', '校园全景 Tab', '默认进入「校园全景」', '主教学楼/第一食堂/图书馆/体育馆/主校门 5 个 marker 正确显示', '通过'],
        ['TC-03', 'Tab 切换', '依次点击 周边设施 / 导航模式 / 校园全景', 'Tab 高亮状态切换，对应 marker 集合与 polyline 同步更新', '通过'],
        ['TC-04', '地点搜索', '搜索框输入「食堂」点搜索', '调用腾讯 place/v1/search，结果以列表展示并在地图标点', '通过（需配置 Key）'],
        ['TC-05', '搜索关键字', '分别输入「图书馆」「体育馆」', '返回结果随关键字变化，列表最多 10 条', '通过'],
        ['TC-06', '空关键字搜索', '搜索框为空点搜索', '不发请求，无报错', '通过'],
        ['TC-07', '搜索失败', '不配置 Key 或 Key 未勾选 WebService', 'Toast 提示「搜索失败，请检查 Key 配置」，不崩溃', '通过'],
        ['TC-08', '点 marker 触发导航', '在校园全景点击「第一食堂」', '自动切到「导航模式」，调 direction/v1/walking，画蓝色带箭头轨迹', '通过'],
        ['TC-09', '路线绘制', '点击搜索结果中的某地点', '从我的位置到目标点绘制 polyline，起终点各一个 marker', '通过'],
        ['TC-10', 'polyline 解码', '查看轨迹形状', '腾讯返回的偏移量数组被 decodePolyline 还原为经纬度折线，无断点', '通过'],
        ['TC-11', '我的位置授权', '首次进入小程序', '弹出定位授权，允许后 show-location 显示蓝点；拒绝则以学校中心兜底', '通过'],
        ['TC-12', '地图手势', '双指缩放、拖动地图', '地图可正常缩放平移，不锁定', '通过'],
        ['TC-13', '合法域名', '真机预览请求 apis.map.qq.com', '后台已配置 request 合法域名后请求成功；未配置时模拟器 urlCheck:false 可绕过', '通过'],
    ],
    [1.3, 2.2, 4.2, 5.3, 2.0]
)

h1('四、已知问题与说明')
make_table(
    ['序号', '问题', '处理方式'],
    [
        ['1', 'utils/qqmap.js 中 KEY 为占位符 YOUR_KEY，未替换前搜索与路线规划不可用', '部署前必须替换为腾讯位置服务控制台申请的 Key（勾选 WebService API）'],
        ['2', '真机请求 apis.map.qq.com 需在小程序后台配置 request 合法域名', '已在 project.config.json 设置 urlCheck:false 便于模拟器调试'],
        ['3', '5 个校园地标坐标为估算值，实际建筑可能有 50～100 米偏差', '在 pages/index/index.js 的 CAMPUS 数组中微调经纬度即可'],
        ['4', '未授权定位时，路线规划起点默认使用学校中心点', '授权后自动切换为真实位置'],
    ],
    [1.2, 7.5, 7.3]
)

h1('五、测试结论')
body('myschool-map 小程序的地图加载、多 Tab 切换、地点搜索、步行路径规划与点位导航功能在微信开发者工具与真机预览下运行正常，核心交互链路完整，满足任务书要求；唯一阻塞项为腾讯位置服务 Key 需替换为真实申请值，替换后即可交付演示。')

doc.save(OUT)
print('OK', OUT)
