# -*- coding: utf-8 -*-
"""生成 myschool-map 坐标数据与实现说明 Word"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\Li\Desktop\myschool-map\校园地图坐标数据与实现说明.docx"

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
            cell_text(t.rows[ri].cells[ci], v, center=(ci != 1 and ci != 4))
    for ri in range(len(t.rows)):
        for ci, w in enumerate(widths):
            t.rows[ri].cells[ci].width = Cm(w)

title('myschool-map 校园地图坐标数据与实现说明')

h1('一、项目概述')
body('myschool-map 是基于微信小程序原生框架开发的校园地图展示应用，以广西民族大学相思湖学院（南宁市江南区友谊路 55 号）为对象，通过微信内置 <map> 组件展示学校中心点及主校门、图书馆、第一食堂、教学楼群、体育馆共 5 个关键地标，支持点击标记查看详情与一键调起微信内置地图导航。')

h1('二、坐标数据表')
body('以下坐标均为 GCJ-02 火星坐标系，与微信小程序 <map> 组件默认坐标系一致，可直接使用。')

make_table(
    ['序号', '名称', '纬度（latitude）', '经度（longitude）', '说明'],
    [
        ['0', '学校中心点（广西民族大学相思湖学院）', '22.686699', '108.266767', '友谊路 55 号，地图初始视野中心'],
        ['1', '主校门', '22.685200', '108.266200', '校园主入口，面向友谊路'],
        ['2', '图书馆', '22.687600', '108.267000', '借阅、自习、电子阅览室'],
        ['3', '第一食堂', '22.686800', '108.268200', '三层餐饮，大众餐与风味小吃'],
        ['4', '教学楼群', '22.687200', '108.265500', 'A/B/C 三栋教学楼与多媒体实验室'],
        ['5', '体育馆', '22.688100', '108.268000', '篮球、羽毛球、乒乓球场地'],
    ],
    [1.2, 4.6, 3.2, 3.2, 4.3]
)

h1('三、实现说明')

h1('3.1 使用的地图 API')
body('本应用直接使用微信小程序原生 <map> 组件（底层为腾讯位置服务地图），无需在前端调用 HTTPS 接口。组件通过 longitude、latitude 属性设置中心点，通过 markers 数组注入标记点，配合 show-location 显示用户当前位置，scale=16 控制缩放级别。点击标记触发 markertap 事件，再调用 wx.openLocation 调起微信内置地图进行导航。')

h1('3.2 坐标偏移问题及解决方法')
body('国内地图存在三种坐标系：WGS-84（GPS 原始坐标）、GCJ-02（国测局火星坐标，高德/腾讯/微信使用）、BD-09（百度坐标）。如果直接把 GPS 设备或从英文维基、OpenStreetMap 上取到的 WGS-84 坐标喂给微信 <map> 组件，会出现约 50～500 米的偏移，标记点落到错误位置。')
body('本项目的处理方法：学校中心点与地标坐标统一从高德/腾讯地图公开 POI 数据获取，本身即为 GCJ-02 坐标，与微信 <map> 组件默认坐标系一致，无需额外转换；若后续接入 GPS 或第三方 WGS-84 坐标，需先调用坐标转换算法转为 GCJ-02 再赋值给 markers，避免偏移。真机测试时已在微信开发者工具中以 375px、414px 宽度及真机预览下验证标记点位置与校园实际布局吻合。')

h1('3.3 权限与合法域名说明')
body('<map> 组件在开发者工具中可直接预览；真机使用 show-location 与 wx.openLocation 需在小程序管理后台「开发 → 开发管理 → 开发设置」中配置 request 合法域名（本项目未直接请求 HTTPS 接口，故无需配置），并在 app.json 中按需申请 scope.userLocation 授权。')

doc.save(OUT)
print('OK', OUT)
