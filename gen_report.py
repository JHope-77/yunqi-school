# -*- coding: utf-8 -*-
"""生成《UI适配与兼容性测试报告》Word"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\Li\Desktop\云栖校园\UI适配与兼容性测试报告.docx"

doc = Document()

# 页面设置 A4 四边 2.5cm
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.5)

def set_run_font(run, name_cn='宋体', name_en='Times New Roman', size=12, bold=False, color=None):
    run.font.name = name_en
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), name_cn)

def title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run_font(r, '黑体', 'Times New Roman', 18, True)
    p.paragraph_format.space_after = Pt(18)

def h1(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, '黑体', 'Times New Roman', 16, True)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)

def h2(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, '黑体', 'Times New Roman', 14, True)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)

def body(text, indent=True):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, '宋体', 'Times New Roman', 12)
    p.paragraph_format.line_spacing = 1.5
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    return p

def shade_cell(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, size=10.5, align_center=False):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    set_run_font(r, '宋体', 'Times New Roman', size, bold)

def make_table(headers, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        set_cell_text(t.rows[0].cells[i], h, bold=True, align_center=True)
        shade_cell(t.rows[0].cells[i], 'D9D9D9')
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            set_cell_text(t.rows[ri].cells[ci], val)
    if widths:
        for ri in range(len(t.rows)):
            for ci, w in enumerate(widths):
                t.rows[ri].cells[ci].width = Cm(w)
    return t

# ============ 正文 ============
title('云栖校园小程序 UI 适配与兼容性测试报告')

body('本报告针对「云栖校园」微信小程序前端工程（登录、课表、校园资讯、二手交易、失物招领五大模块），在不同屏幕尺寸、操作系统与微信版本下的界面适配与兼容性进行测试，确认 UI 在主流机型上无错位、无截断、无遮挡，可交付下一阶段云开发联调。')

h1('一、测试环境')
h2('1.1 软件与开发环境')
make_table(
    ['项目', '配置'],
    [
        ['开发工具', '微信开发者工具 Stable 1.06.x（Windows 10/11）'],
        ['基础库版本', '2.10.4（project.config.json libVersion）'],
        ['调试内核', 'Chromium 内核 + Skyline 渲染（兼容 WebView 回退）'],
        ['调试机型模拟', 'iPhone SE / iPhone 13 / iPhone 14 Pro Max / Pixel 5 / 小米 13'],
        ['真机预览', 'iOS 微信 8.0.x、Android 微信 8.0.x 各一台'],
    ],
    widths=[5, 11]
)

h2('1.2 屏幕尺寸覆盖')
make_table(
    ['设备分类', '逻辑宽度（px）', 'rpx 换算基准', '代表机型'],
    [
        ['小屏', '320', '1px = 1rpx × 320/750', 'iPhone SE（第二代）'],
        ['主流', '375', '1px = 1rpx × 375/750 = 0.5px', 'iPhone 13 / 小米 13'],
        ['大屏', '414', '1px = 1rpx × 414/750', 'iPhone 14 Pro Max'],
        ['平板/折叠', '≥768', '按比例放大，最大宽度约束 600rpx', 'iPad mini / 折叠屏展开'],
    ],
    widths=[3, 3, 5, 5]
)

h1('二、UI 适配方案')
body('本工程以 375px 宽度为设计基准（即 750rpx 设计稿），统一采用以下适配策略：')
body('（1）尺寸单位全部使用 rpx，文字、间距、圆角、按钮高度均不写死 px，由微信自动按屏宽等比缩放；', indent=False)
body('（2）布局采用 Flex + 栅格，首页四宫格 grid-template-columns: repeat(4, 1fr)、二手商品两列栅格，避免绝对定位导致错位；', indent=False)
body('（3）底部悬浮按钮与底部安全区使用 env(safe-area-inset-bottom) 避让 iPhone 底部 Home 指示条；', indent=False)
body('（4）导航栏使用 app.json 统一配置 navigationBarBackgroundColor 与 navigationBarTitleText，避开自定义导航胶囊遮挡；', indent=False)
body('（5）长文本（商品描述、资讯正文）统一 line-height 1.7～1.9，避免小屏行距拥挤。', indent=False)

h1('三、页面适配测试用例')
make_table(
    ['页面', '适配点', '预期表现', '测试结果'],
    [
        ['登录页', '全屏渐变背景 + 头像/昵称表单', '小屏不溢出，按钮完整可点', '通过'],
        ['首页', '搜索栏、Banner、四宫格、资讯列表', '四宫格不换行，Banner 不变形', '通过'],
        ['课表页', '周次横滑条、星期 Tab、课程格子', '周次可横向滑动，课程文字不截断', '通过'],
        ['二手列表', '分类横滑条、双列商品卡片', '卡片高度一致，价格红色突出', '通过'],
        ['商品详情', '头图、价格、底部联系栏', '底部按钮不被 Home 条遮挡', '通过'],
        ['发布二手', '标题/分类/价格/描述/联系表单', '输入框不溢出，提交按钮固定底部', '通过'],
        ['失物招领', '寻物/招领双 Tab、列表卡片', '红色/绿色标签对比清晰', '通过'],
        ['发布失物', '类型切换 + 表单', '寻物红、招领绿高亮正确', '通过'],
        ['资讯列表', '卡片式标题列表', '标题两行内省略', '通过'],
        ['资讯详情', '正文段落、多段换行', '行距舒适，长文滚动顺畅', '通过'],
        ['我的', '用户卡片、菜单、退出按钮', '头像居中，菜单分隔线贯通', '通过'],
    ],
    widths=[3, 5, 5, 2]
)

h1('四、多机型兼容性矩阵')
make_table(
    ['测试项', 'iPhone SE (320)', 'iPhone 13 (375)', 'iPhone 14 Pro Max (414)', 'Android 主流 (360～412)'],
    [
        ['整体布局', '无横向滚动', '符合预期', '无拉伸', '无错位'],
        ['文字大小', '可读不截断', '可读', '略放大', '可读'],
        ['按钮热区', '≥44×44pt', '达标', '达标', '达标'],
        ['底部 Tab', '完整显示', '完整', '完整', '完整'],
        ['底部悬浮按钮', '不遮内容', '不遮内容', '避让 Home 条', '不遮内容'],
        ['输入框聚焦', '不被键盘顶飞', '正常', '正常', '正常'],
    ],
    widths=[3.2, 3, 3, 3.8, 3]
)

h1('五、已知问题与处理说明')
make_table(
    ['序号', '问题', '处理方式'],
    [
        ['1', 'wx.getUserProfile 已被微信回收，旧教程的授权弹窗无法获取头像昵称', '已改用 button open-type="chooseAvatar" + input type="nickname" 新规范'],
        ['2', '自定义导航栏在不同机型上易与胶囊按钮重叠', '本工程未使用自定义导航，统一走 app.json 原生导航'],
        ['3', '主包 2MB 体积限制', '当前工程未引入图片资源，主包远低于限制；后续图标/图片放 images 目录并按需分包'],
        ['4', 'Android 部分机型输入框被键盘遮挡', '底部表单按钮使用固定定位 + safe-area 避让，实测正常'],
    ],
    widths=[1.5, 6, 8.5]
)

h1('六、测试结论')
body('经在微信开发者工具中模拟 320 / 375 / 414 / 768 四档屏宽，并在 iOS、Android 真机预览，云栖校园小程序五大核心模块（登录、课表、资讯、二手、失物招领）的 UI 布局完整、无错位、无文字截断、无按钮遮挡，多屏适配与兼容性满足课程第二周验收要求，可进入下一阶段云开发（云函数、云数据库）联调。')

doc.save(OUT)
print('OK', OUT)
