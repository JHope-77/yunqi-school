# -*- coding: utf-8 -*-
"""生成云栖校园 PRD 与 技术架构/数据库设计 两份 Word 文档"""
import os

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))


def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    def set_font(style, cn, size, bold=False, color=(0, 0, 0)):
        style.font.name = "Times New Roman"
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor(*color)
        rpr = style.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:eastAsia"), cn)
        rf.set(qn("w:ascii"), "Times New Roman")
        rf.set(qn("w:hAnsi"), "Times New Roman")

    set_font(doc.styles["Normal"], "宋体", 12)
    n = doc.styles["Normal"]
    n.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    n.paragraph_format.space_after = Pt(0)
    n.paragraph_format.first_line_indent = Pt(24)
    n.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    set_font(doc.styles["Title"], "黑体", 18, bold=True)
    doc.styles["Title"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.styles["Title"].paragraph_format.space_after = Pt(10)
    doc.styles["Title"].paragraph_format.first_line_indent = Pt(0)

    set_font(doc.styles["Heading 1"], "黑体", 14, bold=True)
    h1 = doc.styles["Heading 1"]
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    h1.paragraph_format.first_line_indent = Pt(0)
    h1.paragraph_format.keep_with_next = True

    set_font(doc.styles["Heading 2"], "黑体", 12, bold=True)
    h2 = doc.styles["Heading 2"]
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    h2.paragraph_format.first_line_indent = Pt(0)
    h2.paragraph_format.keep_with_next = True

    # 页脚页码
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    r.font.size = Pt(9)
    return doc


def body(doc, text, indent=True):
    p = doc.add_paragraph(text)
    if not indent:
        p.paragraph_format.first_line_indent = Pt(0)
    return p


def subtitle(doc, text):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(10)
    for run in p.runs:
        run.font.size = Pt(10.5)
        run.font.name = "Times New Roman"
        run.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), hex_color)
    tc_pr.append(s)


def borders(table):
    pr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:%s" % edge)
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0"); el.set(qn("w:color"), "808080")
        b.append(el)
    pr.append(b)


def repeat_header(row):
    tr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); tr.append(el)


def cant_split(row):
    tr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit"); el.set(qn("w:val"), "true"); tr.append(el)


def set_cell(cell, text, bold=False, center=False, size=10.5):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Times New Roman"
    run.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def make_table(doc, headers, rows, widths, caption=None):
    if caption:
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.first_line_indent = Pt(0)
        cap.paragraph_format.space_before = Pt(4)
        cap.paragraph_format.space_after = Pt(4)
        cap.paragraph_format.keep_with_next = True
        for run in cap.runs:
            run.font.size = Pt(10.5)
            run.font.name = "Times New Roman"
            run.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    borders(table)
    hdr = table.rows[0]
    repeat_header(hdr); cant_split(hdr)
    for i, h in enumerate(headers):
        hdr.cells[i].width = Cm(widths[i])
        shade(hdr.cells[i], "D9D9D9")
        set_cell(hdr.cells[i], h, bold=True, center=True)
    for row_data in rows:
        row = table.add_row()
        cant_split(row)
        for i, v in enumerate(row_data):
            row.cells[i].width = Cm(widths[i])
            set_cell(row.cells[i], str(v), center=(i == 0 and len(str(v)) <= 12))
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


# ============================================================
# 文档一：产品需求文档 PRD
# ============================================================
doc = new_doc()
doc.add_paragraph("云栖校园小程序 产品需求文档（PRD）", style="Title")
subtitle(doc, "版本 V1.0 ｜ 微信小程序 + 云开发 ｜ 编写日期：2026-11-16")

doc.add_heading("一、项目背景与目标", level=1)
body(doc,
    "校园内长期存在信息分散、交易不便的问题：课程表靠班级群转发、闲置物品在 QQ 群里刷屏、失物信息没有固定渠道、"
    "校园活动资讯在多个公众号之间来回切换。本项目面向本校学生，做一个统一入口的微信小程序「云栖校园」，"
    "以微信授权登录为基础，整合课表查询、二手交易、失物招领、校园资讯四个高频场景，V1.0 先跑通"
    "「浏览—发布—联系」的核心闭环，不做在线支付，避免复杂交易流程。")
body(doc,
    "产品目标：让学生在一个小程序内完成每天 80% 的校园轻量信息需求；二手与失物信息按校园场景组织，"
    "发布流程不超过 3 步，列表首屏加载控制在 2 秒内。")

doc.add_heading("二、目标用户与角色", level=1)
make_table(doc,
    ["角色", "描述", "核心诉求", "权限范围"],
    [
        ["普通游客", "未授权访问者", "浏览资讯、查看课表与二手列表", "只读，不能发布"],
        ["在校学生", "微信授权登录，默认可用昵称头像", "发布二手、发布失物、联系同学、看自己的课表", "发布自己的内容，管理自己的发布"],
        ["管理员", "教师或运营同学，后台标记角色", "审核违规内容、下架商品、置顶资讯", "全部内容可读可写，可封禁账号"],
    ],
    [2.2, 4.6, 4.6, 4.1],
    caption="表1  用户角色与权限")

doc.add_heading("三、使用场景", level=1)
make_table(doc,
    ["编号", "场景", "用户故事", "优先级"],
    [
        ["S1", "查课表", "早上出门前想知道今天第二节课在哪上，打开小程序直接看当天课表", "P0"],
        ["S2", "出二手", "毕业季教材太多，拍照填写价格后发布，等同学来联系", "P0"],
        ["S3", "买二手", "开学想买一本二手教材，按分类筛选，看上后联系卖家", "P0"],
        ["S4", "寻物/招领", "U盘丢在图书馆，发一条寻物启事；捡到校园卡，发一条招领启事", "P0"],
        ["S5", "看资讯", "查看学校讲座、社团活动、通知公告", "P0"],
        ["S6", "跑腿代办", "帮同学带饭、取快递，发布或接单（V1.1 预留）", "P1"],
    ],
    [1.4, 2.4, 8.6, 1.6],
    caption="表2  核心使用场景")

doc.add_heading("四、功能清单（V1.0）", level=1)
make_table(doc,
    ["模块", "功能点", "说明", "优先级"],
    [
        ["首页", "搜索 + 快捷入口 + 资讯推荐", "顶部搜索框，宫格进入各模块，首页推荐 3 条最新资讯", "P0"],
        ["课表", "周次切换 + 课程格子", "按周一至周五、第 1-10 节显示课程，显示教室与教师", "P0"],
        ["二手交易", "商品列表", "分类筛选 + 关键词搜索 + 卡片列表", "P0"],
        ["二手交易", "商品详情", "大图、价格、描述、卖家信息、联系按钮", "P0"],
        ["二手交易", "发布商品", "标题、分类、价格、描述、最多 3 张图、联系方式", "P0"],
        ["二手交易", "我的发布", "查看自己发布的商品，可下架、标记已售出", "P0"],
        ["失物招领", "双 Tab 列表", "「寻物启事」与「招领启事」切换，按时间倒序", "P0"],
        ["失物招领", "发布启事", "物品名称、地点、时间、描述、联系方式", "P0"],
        ["校园资讯", "资讯列表 + 详情", "标题、封面、摘要、发布时间；详情页分段排版", "P0"],
        ["个人中心", "授权登录 + 我的入口", "头像昵称、我的发布、我的收藏、设置", "P0"],
        ["跑腿代办", "发布与接单", "定位学校、取送时间、费用", "P1"],
        ["消息通知", "审核结果、下架提醒", "订阅消息推送", "P1"],
    ],
    [2.2, 3.0, 7.3, 1.4],
    caption="表3  V1.0 功能清单与优先级")
body(doc,
    "V2.0 暂不实现：在线支付与担保交易、同学互评信用分、食堂评价、校园地图导航、积分签到体系，"
    "避免 V1.0 范围失控。")

doc.add_heading("五、页面结构", level=1)
body(doc,
    "小程序采用底部标签导航，共 5 个一级 Tab：首页、课表、二手、失物、我的。"
    "二级页面从一级 Tab 内进入：商品详情、发布二手、资讯详情、发布失物、我的发布等。")
make_table(doc,
    ["一级 Tab", "包含页面", "关键元素"],
    [
        ["首页", "index", "搜索栏、5 宫格入口、Banner、资讯推荐列表"],
        ["课表", "schedule", "周次选择器、星期切换、课程格子（按节次排列）"],
        ["二手", "market、goods-detail、publish-goods", "分类筛选、商品卡片、详情、发布表单"],
        ["失物", "lost、publish-lost", "寻物/招领双 Tab、信息卡片、发布表单"],
        ["我的", "profile", "用户信息卡、我的发布、我的收藏、设置"],
    ],
    [2.0, 5.6, 6.3],
    caption="表4  页面与 Tab 结构")

doc.add_heading("六、关键业务流程", level=1)
doc.add_heading("6.1  二手商品发布与交易", level=2)
body(doc,
    "学生在二手 Tab 点击发布 → 选择分类 → 填写标题与价格（必填）→ 上传最多 3 张图 → 填写联系方式 → 提交。"
    "云函数校验必填项后写入 posts 集合，默认状态为在售。买家在列表或详情点击「联系卖家」，"
    "通过小程序客服消息或展示微信号与卖家沟通。卖家在「我的发布」中可将商品标记为已售出或下架。"
    "整个流程不接入微信支付，避免担保交易的合规与资金问题。")
doc.add_heading("6.2  失物招领发布", level=2)
body(doc,
    "用户在失物 Tab 先选择「寻物启事」或「招领启事」，再点击发布：填写物品名称、丢失/捡到地点、时间、"
    "联系方式，可上传一张图片。提交后出现在对应 Tab 列表中；物品找回后发布者可自行关闭该启事。")
doc.add_heading("6.3  资讯浏览", level=2)
body(doc,
    "管理员在云后台或通过管理员账号发布资讯（标题、封面、摘要、正文），学生在首页推荐与资讯 Tab 看到列表，"
    "点击进入详情页，正文按段落排版。置顶资讯排在列表最前。")

doc.add_heading("七、非功能需求", level=1)
make_table(doc,
    ["类别", "要求"],
    [
        ["性能", "列表首屏渲染 ≤2 秒；小程序主包 ≤2MB；云函数冷启动 ≤1.5 秒"],
        ["安全", "云数据库按「创建者可读可写、他人只读」授权；敏感词过滤；手机号等联系方式在云函数内脱敏返回"],
        ["合规", "不收集身份证、银行卡等敏感信息；用户协议与隐私政策入口在「我的—设置」"],
        ["兼容性", "微信基础库 ≥2.10，适配主流 iOS/Android 机型，最低屏幕 320px"],
        ["可用性", "列表为空、网络失败、未登录三种状态都有明确引导，不允许白屏"],
    ],
    [2.2, 13.3],
    caption="表5  非功能需求")

doc.add_heading("八、里程碑与分工建议", level=1)
make_table(doc,
    ["时间", "里程碑", "产出"],
    [
        ["第 1-2 天", "用户调研与竞品分析", "访谈记录、竞品对比表、功能清单确认"],
        ["第 3 天", "技术选型与数据库设计", "架构图、JSON 集合设计文档"],
        ["第 4-5 天", "原型与 UI 规范", "高保真原型、UI 色值/字号规范"],
        ["第 6-7 天", "云环境初始化与归档", "云开发环境可用、Git 仓库、交付物打包"],
    ],
    [2.4, 5.6, 7.5],
    caption="表6  一周排期")

out1 = os.path.join(BASE, "云栖校园产品需求文档PRD.docx")
doc.save(out1)
print("saved:", out1)


# ============================================================
# 文档二：技术架构与数据库设计文档
# ============================================================
doc = new_doc()
doc.add_paragraph("云栖校园小程序 技术架构与数据库设计文档", style="Title")
subtitle(doc, "微信原生小程序 + 云开发 Serverless ｜ 版本 V1.0")

doc.add_heading("一、技术选型", level=1)
body(doc,
    "项目采用「微信原生小程序 + 云开发 Serverless」架构：前端不写传统后端，业务逻辑放在云函数，"
    "数据存云数据库（JSON 文档型），图片存云存储。选型理由是免域名备案、免运维服务器、"
    "与微信登录天然打通，适合课程演示与小型校园应用快速上线。")
make_table(doc,
    ["层次", "技术", "版本/说明"],
    [
        ["前端", "微信原生框架 WXML / WXSS / JS + WeUI", "基础库 ≥2.10，不使用第三方框架"],
        ["后端", "微信云开发云函数", "Node.js 16，按业务模块拆分函数"],
        ["数据库", "云数据库（JSON 文档型）", "集合式存储，自带 _openid 与权限规则"],
        ["文件存储", "云存储", "商品图、失物图、资讯封面"],
        ["登录", "wx.login + 云函数 openid 校验", "免密获取用户唯一标识"],
        ["协作", "Git + 代码规范", "main 分支保护，功能分支合并"],
    ],
    [2.2, 5.8, 7.5],
    caption="表1  技术栈")

doc.add_heading("二、系统架构", level=1)
body(doc,
    "整体分为三层：客户端视图层负责渲染与交互；云函数层承接所有业务逻辑与权限校验；"
    "数据层由云数据库与云存储组成。客户端不直接读写数据库，所有写操作都经过云函数，"
    "避免把密钥与规则写在前端。架构关系如下图所示（同目录下另附「云栖校园系统架构图.drawio」可直接打开）。")
make_table(doc,
    ["层级", "组成", "职责"],
    [
        ["视图层", "小程序 5 个 Tab 页面 + 二级页", "展示数据、表单收集、调用云函数"],
        ["云函数层", "login、market、lost、news、schedule、upload", "登录态、列表查询、发布校验、下架、脱敏返回"],
        ["数据层", "users、posts、lost_items、news、schedule 集合 + 云存储", "持久化业务数据与图片文件"],
    ],
    [2.0, 6.6, 6.9],
    caption="表2  分层职责")

doc.add_heading("三、云数据库集合设计", level=1)
body(doc, "数据库为 JSON 文档型，每个集合一行就是一条文档。下面列出 V1.0 的核心集合、字段、索引与权限。")

doc.add_heading("3.1  users（用户）", level=2)
make_table(doc,
    ["字段", "类型", "说明", "索引"],
    [
        ["_id", "string", "文档主键，系统自动生成", "主键"],
        ["_openid", "string", "微信用户唯一标识，系统自动写入", "唯一索引"],
        ["nickName", "string", "微信昵称，授权后写入", "-"],
        ["avatar", "string", "头像 fileID", "-"],
        ["role", "string", "user / admin，默认 user", "-"],
        ["status", "string", "normal / banned", "-"],
        ["createdAt", "Date", "首次登录时间", "普通索引"],
    ],
    [2.6, 2.0, 7.4, 3.5],
    caption="表3  users 集合")

doc.add_heading("3.2  posts（二手商品）", level=2)
make_table(doc,
    ["字段", "类型", "说明", "索引"],
    [
        ["_id / _openid", "string", "主键 / 发布者", "_openid"],
        ["title", "string", "标题，必填，≤40 字", "-"],
        ["desc", "string", "商品描述，≤500 字", "-"],
        ["price", "number", "价格，单位元", "-"],
        ["category", "string", "教材 / 电子 / 生活用品 / 代步 / 其他", "category"],
        ["images", "array<string>", "云存储 fileID，最多 3 张", "-"],
        ["location", "string", "交易地点，如「东区宿舍」", "-"],
        ["contact", "string", "联系方式（云函数脱敏）", "-"],
        ["status", "string", "on / sold / off", "status"],
        ["views", "number", "浏览数，详情页 +1", "-"],
        ["createdAt", "Date", "发布时间", "createdAt 倒序"],
    ],
    [2.6, 2.2, 7.2, 3.5],
    caption="表4  posts 集合")

doc.add_heading("3.3  lost_items（失物招领）", level=2)
make_table(doc,
    ["字段", "类型", "说明", "索引"],
    [
        ["_id / _openid", "string", "主键 / 发布者", "-"],
        ["type", "string", "miss（寻物） / found（招领）", "type"],
        ["title", "string", "物品名称", "-"],
        ["desc", "string", "特征描述", "-"],
        ["images", "array<string>", "图片 fileID", "-"],
        ["place", "string", "丢失/捡到地点", "-"],
        ["time", "string", "发生时间", "-"],
        ["contact", "string", "联系方式", "-"],
        ["status", "string", "open / closed", "status"],
        ["createdAt", "Date", "发布时间", "createdAt 倒序"],
    ],
    [2.6, 2.2, 7.2, 3.5],
    caption="表5  lost_items 集合")

doc.add_heading("3.4  news（校园资讯）", level=2)
make_table(doc,
    ["字段", "类型", "说明", "索引"],
    [
        ["title", "string", "资讯标题", "-"],
        ["summary", "string", "列表摘要，≤80 字", "-"],
        ["content", "string", "正文，分段存储", "-"],
        ["cover", "string", "封面图 fileID", "-"],
        ["publisher", "string", "发布者 openid", "-"],
        ["isTop", "boolean", "是否置顶", "-"],
        ["views", "number", "浏览数", "-"],
        ["createdAt", "Date", "发布时间", "isTop 优先，再按 createdAt"],
    ],
    [2.6, 2.2, 7.2, 3.5],
    caption="表6  news 集合")

doc.add_heading("3.5  schedule（课表）", level=2)
make_table(doc,
    ["字段", "类型", "说明", "索引"],
    [
        ["_openid", "string", "所属学生", "_openid"],
        ["day", "number", "星期 1-5", "-"],
        ["section", "number", "节次 1-10", "-"],
        ["courseName", "string", "课程名", "-"],
        ["teacher", "string", "教师", "-"],
        ["room", "string", "教室", "-"],
        ["weeks", "array<number>", "上课周次，如 [1,2,3,4,5,6,7,8]", "-"],
    ],
    [2.6, 2.4, 7.0, 3.5],
    caption="表7  schedule 集合")

doc.add_heading("3.6  集合权限规则", level=1)
make_table(doc,
    ["集合", "读", "写", "说明"],
    [
        ["users", "仅本人", "仅本人（role 字段仅管理员可改）", "openid 自动匹配"],
        ["posts", "所有人", "仅创建者（status 可自行修改）", "下架与售出由发布者操作"],
        ["lost_items", "所有人", "仅创建者", "关闭启事由发布者操作"],
        ["news", "所有人", "仅 role=admin", "学生无发布权限"],
        ["schedule", "仅本人", "仅本人", "V1.0 先手动录入"],
    ],
    [2.4, 3.2, 5.6, 4.3],
    caption="表8  云数据库权限规则")

doc.add_heading("四、云函数 API 清单", level=1)
body(doc,
    "云函数以 callFunction 方式调用，每个函数内部按 action 分发。所有写操作先做登录态与必填项校验，"
    "联系方式在返回前做中间四位脱敏。")
make_table(doc,
    ["云函数", "action", "入参", "出参"],
    [
        ["login", "login", "无", "openid、role、是否新用户"],
        ["market", "list", "category、keyword、page、size", "posts 分页列表"],
        ["market", "detail", "id", "单条商品详情，浏览数 +1"],
        ["market", "publish", "title/price/category/desc/images/contact", "新记录 _id"],
        ["market", "setStatus", "id、status(on/sold/off)", "成功/失败"],
        ["lost", "list", "type、page", "失物启事列表"],
        ["lost", "publish", "type/title/place/time/contact/images", "新记录 _id"],
        ["news", "list", "page、size", "资讯列表（置顶在前）"],
        ["news", "detail", "id", "资讯详情，浏览数 +1"],
        ["schedule", "get", "week", "本人本周课程"],
        ["schedule", "save", "课程数组", "保存结果"],
        ["upload", "file", "临时文件路径", "云存储 fileID，替换后落库"],
    ],
    [2.0, 2.0, 6.0, 5.5],
    caption="表9  云函数接口清单")

doc.add_heading("五、云存储规划", level=1)
make_table(doc,
    ["目录", "内容", "权限"],
    [
        ["/users/avatar/{openid}/", "用户头像", "公开读、仅本人写"],
        ["/posts/{openid}/", "二手商品图（最多 3 张）", "公开读、仅本人写"],
        ["/lost/{openid}/", "失物招领图", "公开读、仅本人写"],
        ["/news/cover/", "资讯封面", "公开读、仅管理员写"],
    ],
    [4.4, 5.6, 5.5],
    caption="表10  云存储目录规划")
body(doc,
    "图片上传采用「先选图 → 云存储直传换 fileID → 提交业务表单」两步走，避免把大文件塞进云函数请求体。"
    "列表缩略图在详情页才加载原图，减轻首屏压力。")

doc.add_heading("六、前端目录结构", level=1)
body(doc,
    "项目根目录下 pages 存放 5 个 Tab 与二级页面，components 存放可复用卡片组件，"
    "utils 封装云函数调用与登录态判断，cloudfunctions 与 pages 同级，按模块拆分函数。", indent=False)
make_table(doc,
    ["目录", "内容"],
    [
        ["pages/index、schedule、market、lost、profile", "五个一级 Tab 页面"],
        ["pages/goods-detail、publish-goods、publish-lost、news-detail", "二级页面"],
        ["components/post-card、lost-card、news-item", "可复用列表卡片"],
        ["utils/cloud.js", "callFunction 封装、登录态检查、错误提示"],
        ["cloudfunctions/market、lost、news、schedule、login、upload", "云函数源码"],
    ],
    [6.6, 8.9],
    caption="表11  前端目录约定")

doc.add_heading("七、安全与性能考虑", level=1)
body(doc,
    "安全上：写操作必须经过云函数校验，前端不直接操作数据库；联系方式在出参脱敏；"
    "管理员角色不可由前端自助修改；敏感词在发布时过滤。性能上：列表分页每页 10 条，"
    "按 createdAt 倒序建索引；图片上传后生成缩略图；云函数冷启动期间用骨架屏占位。"
    "V1.0 不引入 WebSocket 与消息推送，待 V1.1 再接入订阅消息。")

out2 = os.path.join(BASE, "云栖校园技术架构与数据库设计文档.docx")
doc.save(out2)
print("saved:", out2)
