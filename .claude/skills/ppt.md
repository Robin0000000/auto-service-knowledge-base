---
name: ppt
description: 当用户要求做PPT、生成PPT、制作PPT演示文稿时使用。用 python-pptx 生成专业宽屏(16:9)演示文稿，配色统一、结构清晰。
allowed-tools: [Bash, Write, Read, Edit, Grep, Glob]
---

# PPT 生成技能

你是一个专业的演示文稿设计师。当用户要求生成 PPT 时，按照以下流程操作。

## 触发条件

用户说「做PPT」「生成PPT」「制作PPT」「写个PPT」「/ppt」或明确要求生成演示文稿时触发。

## 核心原则

1. **宽屏比例**：统一使用 16:9 (`slide_width=13333333`, `slide_height=7500000` EMU)
2. **配色统一**：主色深蓝 (#1A56DB)，强调色浅蓝 (#2591CD)，深色正文 (#1E293B)，浅灰背景 (#F0F4FA)
3. **结构清晰**：每页有明确主题，顶部 header bar + 卡片式内容布局
4. **中文字体**：优先使用 Microsoft YaHei（微软雅黑），fallback SimHei
5. **输出格式**：`.pptx` 文件，路径由用户指定，默认 `docs/演示文稿.pptx`

## 执行流程

### Step 1：确认需求

如果用户没有明确说明，用 AskUserQuestion 确认：
- PPT 主题/标题
- 页数和每页内容要点
- 输出文件路径

如果用户已说了大致内容（如「介绍XX项目，5页」），则先规划每页的内容结构，展示给用户确认后再生成。

### Step 2：编写生成脚本

编写一个临时 Python 脚本（`_generate_ppt.py`），使用 `python-pptx` 库。脚本结构：

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9 宽屏
prs.slide_height = Inches(7.5)

# 配色常量
C_PRIMARY   = RGBColor(0x1A, 0x56, 0xDB)   # 深蓝
C_ACCENT    = RGBColor(0x25, 0x91, 0xCD)   # 浅蓝
C_ORANGE    = RGBColor(0xE8, 0x6A, 0x17)   # 橙色强调
C_GREEN     = RGBColor(0x1B, 0x9E, 0x4B)   # 绿色
C_RED       = RGBColor(0xDC, 0x35, 0x35)   # 红色
C_DARK      = RGBColor(0x1E, 0x29, 0x3B)   # 深色正文
C_GRAY      = RGBColor(0x6B, 0x72, 0x80)   # 灰色辅助
C_BG_LIGHT  = RGBColor(0xF0, 0xF4, 0xFA)   # 浅蓝灰背景
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
FONT        = "Microsoft YaHei"

def add_bg(slide, color=C_WHITE):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text(slide, left, top, width, height, text, size=14, color=C_DARK, bold=False, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = FONT
    p.alignment = align
    return txBox

def add_header(slide, title, subtitle=""):
    """统一的页面顶部栏"""
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(1.1), C_PRIMARY)
    add_text(slide, Inches(0.8), Inches(0.18), Inches(10), Inches(0.5),
             title, size=30, color=C_WHITE, bold=True)
    if subtitle:
        add_text(slide, Inches(0.8), Inches(0.65), Inches(10), Inches(0.35),
                 subtitle, size=14, color=RGBColor(0xB0, 0xCC, 0xF0))

def add_card(slide, left, top, width, height, title, items, title_color=C_PRIMARY, bar_color=C_ACCENT):
    """卡片：圆角矩形背景 + 顶部色条 + 标题 + 条目列表"""
    add_rect(slide, left, top, width, height, C_WHITE)
    add_rect(slide, left, top, width, Inches(0.06), bar_color)
    add_text(slide, left+Inches(0.25), top+Inches(0.15), width-Inches(0.5), Inches(0.4),
             title, size=16, color=title_color, bold=True)
    y = top + Inches(0.6)
    for icon, text in items:
        add_text(slide, left+Inches(0.25), y, width-Inches(0.5), Inches(0.32),
                 f"{icon}  {text}", size=11.5, color=C_DARK)
        y += Inches(0.3)

# === 在此编写每页幻灯片 ===
# ...

prs.save(output_path)
print(f"PPT已生成：{output_path}")
```

### Step 3：执行并验证

```bash
cd <项目根目录> && python _generate_ppt.py
```

验证生成成功：确认文件存在且大小合理（通常 30-100KB）。

### Step 4：清理

删除临时脚本 `_generate_ppt.py`，告知用户 PPT 文件的路径。

## 常用布局模式

### 标题页 (Title Slide)
- 全幅深蓝背景 + 居中白色大标题 + 副标题 + 日期/作者

### 内容页 (Content Slide)
- 顶部深蓝 header bar (标题+副标题)
- 主体：左右分栏或 2×2 / 3 列卡片布局
- 每张卡片：顶部彩色细条 + 标题 + 条目列表

### 流程图页 (Flow Slide)
- 横向 4-5 步骤卡片，用箭头 (▶) 连接
- 每步卡片：顶部色条 + 步骤名 + 简要说明

### 对比页 (Comparison Slide)
- 左右两栏：左侧「优势/正确做法」(绿色调)，右侧「劣势/错误做法」(红色调)

### 总结页 (Summary Slide)
- 全幅深蓝背景 + 居中核心观点 + 要点列表

## 注意事项

1. **先检查 python-pptx 是否可用**：`python -c "import pptx"`，不可用则 `pip install python-pptx`（首次使用需用户授权）
2. **内容适配页面**：每页不宜超过 5-6 个要点，文字过多时分页
3. **保持一致性**：全 PPT 使用同一套配色和字体，标题/正文字号统一
4. **使用环境已有数据**：如果用户要求介绍某个代码项目，先分析项目代码提取真实内容
5. **脚本执行完立即删除**：`_generate_ppt.py` 是临时文件，生成后清理
