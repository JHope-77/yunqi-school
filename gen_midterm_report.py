# -*- coding: utf-8 -*-
"""生成《云栖校园》中期报告（基于实际项目进度填充模板）"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\Li\Desktop\云栖校园\云栖校园中期报告.docx"

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
    p.paragraph_format.space_after = Pt(12)

def h1(t):
    p = doc.add_paragraph()
    set_run(p.add_run(t), '黑体', 'Times New Roman', 15, True)
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(6)

def h2(t):
    p = doc.add_paragraph()
    set_run(p.add_run(t), '黑体', 'Times New Roman', 13, True)
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(4)

def body(t, indent=True):
    p = doc.add_paragraph()
    set_run(p.add_run(t), '宋体', 'Times New Roman', 12)
    p.paragraph_format.line_spacing = 1.5
    if indent: p.paragraph_format.first_line_indent = Pt(24)

def shade(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def cell_text(cell, text, bold=False, center=False, size=10.5):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    set_run(p.add_run(text), '宋体', 'Times New Roman', size, bold)

def make_table(headers, rows, widths):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        cell_text(t.rows[0].cells[i], h, bold=True, center=True); shade(t.rows[0].cells[i])
    for ri, row in enumerate(rows, start=1):
        for ci, v in enumerate(row):
            cell_text(t.rows[ri].cells[ci], v, center=(ci == len(headers)-1))
    for ri in range(len(t.rows)):
        for ci, w in enumerate(widths):
            t.rows[ri].cells[ci].width = Cm(w)

# ============ 内容 ============
title('《云栖校园》小程序开发中期报告')

p = doc.add_paragraph()
set_run(p.add_run('项目名称：云栖校园——一站式智慧校园生活服务平台小程序开发'), size=12)
p = doc.add_paragraph()
set_run(p.add_run('学生姓名：[您的姓名]    学号：[您的学号]'), size=12)
p = doc.add_paragraph()
set_run(p.add_run('所属子项目：前端开发与核心模块实现'), size=12)
p = doc.add_paragraph()
set_run(p.add_run('报告日期：2026 年 11 月 16 日'), size=12)

h1('一、项目概述与需求分析')
h2('1.1 项目背景')
body('本项目旨在开发一款基于微信小程序的"云栖校园"平台，解决高校师生信息碎片化痛点。通过集成课表查询、校园资讯、二手交易、失物招领等功能，打造便捷高效的智慧校园数字生活助手。')

h2('1.2 核心功能模块梳理')
body('根据需求调研，本项目主要包含以下核心模块：')
body('首页：聚合入口，顶部搜索栏、Banner、四宫格功能入口、校园资讯推荐列表；', indent=False)
body('课表查询：周次横滑选择、星期 Tab 切换，按大节展示课程名称、教室、教师；', indent=False)
body('二手交易：分类筛选列表、商品详情、发布表单（标题/分类/价格/描述/联系方式）；', indent=False)
body('失物招领：寻物/招领双 Tab 列表、发布表单；', indent=False)
body('校园资讯：公告列表与详情正文；', indent=False)
body('个人中心：微信头像昵称授权登录、我的发布、我的收藏、退出登录。', indent=False)

h1('二、原型设计与交互逻辑')
h2('2.1 原型设计说明')
body('设计工具：HTML 高保真交互原型（任务书允许 Axure 源文件 或 HTML 导出包二选一）；')
body('设计思路：采用底部 TabBar 导航结构（首页/课表/二手/失物/我的），确保核心功能一键直达。页面主色 #2E6BFF，卡片圆角 20rpx，统一使用 rpx 单位做屏幕适配；关键按钮置于底部安全区内，遵循移动端拇指热区原则。')

h2('2.2 原型图展示')
body('（此处插入原型截图：首页、课表、二手列表、商品详情、失物招领、个人中心）')
body('图 1：云栖校园小程序首页原型图', indent=False)
body('图 2：二手交易详情页原型图', indent=False)

h1('三、前端开发实施情况（WXML/WXSS）')
h2('3.1 技术栈说明')
body('开发工具：微信开发者工具 Stable 1.06.x（Windows 11）；')
body('核心技术：微信小程序原生框架（WXML/WXSS/JS）；')
body('UI 组件：基于 WeUI 设计规范自建全局样式类（weui-cell 风格的卡片、按钮、标签、搜索栏），主色与圆角统一在 app.wxss 中管理；')
body('数据层：当前阶段使用 utils/mock.js 提供 Mock 数据，字段与后续云数据库集合对齐，便于下一阶段无缝切换 wx.cloud.callFunction。')

h2('3.2 页面开发进度')
make_table(
    ['页面名称', '开发状态', '说明'],
    [
        ['登录页 (login)', '已完成', 'chooseAvatar + nickname 新规范授权，未登录自动跳转'],
        ['首页 (index)', '已完成', '搜索栏、Banner、四宫格、资讯推荐列表'],
        ['课表页 (schedule)', '已完成', '周次横滑、星期 Tab、5 大节课程格子'],
        ['二手列表 (market)', '已完成', '分类 Chips、双列商品卡片、右下角发布悬浮按钮'],
        ['商品详情 (goods-detail)', '已完成', '头图、价格、卖家信息、底部联系栏'],
        ['发布二手 (publish-goods)', '已完成', '标题/分类/价格/描述/联系表单，含校验'],
        ['失物招领 (lost)', '已完成', '寻物/招领双 Tab、卡片式列表'],
        ['发布失物 (publish-lost)', '已完成', '类型切换 + 地点/时间/描述表单'],
        ['校园资讯 (news)', '已完成', '公告列表与详情正文'],
        ['个人中心 (profile)', '已完成', '用户卡片、菜单、退出登录二次确认'],
    ],
    [4.5, 2.5, 9.0]
)

h2('3.3 页面运行截图展示')
body('（此处插入微信开发者工具模拟器截图：首页、二手列表、课表页在 iPhone SE 与 iPhone 14 Pro Max 两档尺寸下的对比）')
body('图 3：首页页面运行效果截图', indent=False)
body('图 4：二手列表页页面运行效果截图', indent=False)
body('图 5：课表页多尺寸适配截图（iPhone SE 与 iPhone 14 Pro Max）', indent=False)

h2('3.4 关键代码实现')
body('（1）WeUI 风格组件应用：在二手列表、失物招领等页面使用统一的 .card 卡片、.tag 标签、.btn-primary 按钮类，配合 Flex + Grid 布局快速实现图文混排；')
body('（2）响应式布局：全工程使用 rpx 单位（750rpx 设计稿），底部悬浮按钮使用 env(safe-area-inset-bottom) 避让 iPhone Home 指示条；')
body('（3）请求封装：utils/request.js 统一返回 Promise，当前返回 Mock 数据，注释中预留 wx.cloud.callFunction 替换位置。')

h1('四、存在问题与整改计划')
h2('4.1 当前存在的问题')
body('（1）样式细节：部分长文本在小屏（320px）下行距偏紧，需在真机调试中进一步微调；', indent=False)
body('（2）图片资源：当前使用渐变色块占位，后续需替换为真实校园图片并做懒加载；', indent=False)
body('（3）交互逻辑：Mock 数据已在前端跑通，但尚未接入真实云数据库，数据刷新与收藏、发布等写操作需等待后端联调。', indent=False)

h2('4.2 整改计划')
body('（1）样式偏差：在开发者工具中切换 iPhone SE（320）、iPhone 13（375）、iPhone 14 Pro Max（414）三档尺寸逐页检查，微调 margin/padding；', indent=False)
body('（2）真实图片：使用云存储存放 Banner 与商品图，通过 fileID 引用；', indent=False)
body('（3）代码规范：统一 WXML 标签闭合、清理未使用的 WXSS 类、为关键函数补充注释。', indent=False)

h1('五、后续工作计划')
h2('5.1 下一阶段目标')
body('（1）云开发接入：创建 users / posts / lost_items / news / schedule 五个云数据库集合，编写 login、market、lost、news、schedule 等云函数，实现增删改查；', indent=False)
body('（2）腾讯位置服务 API：在 myschool-map 子模块中已完成地图组件、地点搜索与步行路径规划，下一阶段接入真实 Key 并与主工程合并；', indent=False)
body('（3）联调与真机测试：完成前后端数据联调，在 iOS / Android 真机上验证登录、发布、收藏等核心流程。', indent=False)

h2('5.2 时间节点')
body('第 8 周：完成云数据库设计与云函数编写；', indent=False)
body('第 9 周：前后端联调与 Bug 修复；', indent=False)
body('第 10 周：全链路真机测试、UI 走查与答辩准备。', indent=False)

h1('附件')
body('原型设计链接：[粘贴 HTML 原型文件路径或链接]', indent=False)
body('代码仓库地址：[粘贴 Git 仓库链接]', indent=False)

doc.save(OUT)
print('OK', OUT)
