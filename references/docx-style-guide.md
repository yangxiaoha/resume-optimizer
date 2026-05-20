# Word文档样式规范 / DOCX Style Guide

本文件定义简历优化Skill输出Word文档的完整样式规范与可复用代码模板。

---

## 颜色体系

| 用途 | RGB值 | 场景 |
|------|-------|------|
| DARK_BLUE | (0x1A, 0x3A, 0x5C) | 姓名、公司名、项目标题 |
| MID_BLUE | (0x2E, 0x75, 0xB6) | 章节标题、职位名 |
| GRAY | (0x60, 0x60, 0x60) | 时间、副标题、技术栈 |
| BLACK | (0x1A, 0x1A, 0x1A) | 正文内容 |
| ACCENT | (0x0D, 0x6E, 0x6E) | 标签前缀粗体（【架构设计】等） |

---

## 页面设置

```python
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
```

---

## 通用工具函数

```python
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DARK_BLUE = RGBColor(0x1A, 0x3A, 0x5C)
MID_BLUE  = RGBColor(0x2E, 0x75, 0xB6)
GRAY      = RGBColor(0x60, 0x60, 0x60)
BLACK     = RGBColor(0x1A, 0x1A, 0x1A)
ACCENT    = RGBColor(0x0D, 0x6E, 0x6E)

def add_para_border_bottom(paragraph, color_hex="2E75B6"):
    """为段落添加下边框线（用于章节标题）"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def section_heading(doc, title):
    """章节标题：中蓝12pt粗体 + 下边框"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = MID_BLUE
    add_para_border_bottom(p)

def job_header(doc, company, title, period):
    """工作经历标题：公司名+时间段 / 职位名"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(1)
    r1 = p.add_run(company)
    r1.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = DARK_BLUE
    r1b = p.add_run("    " + period)
    r1b.font.size = Pt(10)
    r1b.font.color.rgb = GRAY
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(3)
    r2 = p2.add_run(title)
    r2.italic = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = MID_BLUE

def bullet(doc, text, bold_prefix=None, indent=0.5):
    """正文子弹点：可选粗体前缀标签"""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(indent)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.bold = True
        rb.font.size = Pt(10)
        rb.font.color.rgb = ACCENT
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = BLACK

def proj_title(doc, title):
    """项目标题：深蓝10.5pt粗体，带▶前缀"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("▶ " + title)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK_BLUE

def proj_meta(doc, text):
    """项目技术栈行：灰色斜体，缩进"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = GRAY
    r.italic = True

def proj_bullet(doc, text, bold_prefix=None):
    """项目内子弹点：缩进1cm"""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(1.0)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.bold = True
        rb.font.size = Pt(10)
        rb.font.color.rgb = ACCENT
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = BLACK
```

---

## 文档结构模板

```python
def build_resume(candidate, jd_target):
    doc = Document()
    
    # 页面设置
    for section in doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # ① 姓名（24pt深蓝居中）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(candidate['name'])
    r.bold = True; r.font.size = Pt(24); r.font.color.rgb = DARK_BLUE

    # ② 联系信息（10pt灰色居中）
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"{candidate['phone']}   |   {candidate['email']}   |   {candidate['location']}")
    r2.font.size = Pt(10); r2.font.color.rgb = GRAY

    # ③ 定位副标题（10pt中蓝居中粗体 + 下划线）
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run(f"{jd_target['role_direction']}  |  {candidate['exp_years']}年Java开发  |  {candidate['highlights']}")
    r3.font.size = Pt(10); r3.font.color.rgb = MID_BLUE; r3.bold = True
    add_para_border_bottom(p3, "CCCCCC")

    # ④ 基本信息
    section_heading(doc, "基本信息 / Basic Info")
    # ...

    # ⑤ 个人优势
    section_heading(doc, "个人优势 / Profile")
    # ...

    # ⑥ 专业技能（按JD要求顺序排列）
    section_heading(doc, "专业技能 / Technical Skills")
    # ...

    # ⑦ 工作经历（含【标签】前缀）
    section_heading(doc, "工作经历 / Work Experience")
    # ...

    # ⑧ 核心项目（含★标注和对标注释）
    section_heading(doc, "核心项目经验 / Key Projects")
    # ...

    # ⑨ 发明专利（如有）
    # section_heading(doc, "发明专利 / Patents")

    # ⑩ 教育背景
    section_heading(doc, "教育背景 / Education")
    # ...

    return doc
```

---

## 注意事项

1. **字符串中的中文引号**：Python字符串中避免混用中文双引号`""`，改用单引号包裹字符串，或将中文引号替换为`「」`
2. **特殊字符**：箭头`→`、特殊符号可能引发Python语法错误，建议用`-`或`>`替代
3. **输出路径**：统一输出到 `/mnt/user-data/outputs/姓名_优化简历_目标岗位.docx`
4. **验证**：生成后检查文件是否存在，用 `present_files` 工具呈现给用户
