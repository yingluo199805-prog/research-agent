# Word文档生成指南

当用户要求生成.docx格式纪要时，执行以下流程。

## 依赖检查

```bash
pip install python-docx --break-system-packages
```

## 标准生成脚本

```python
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

def create_meeting_notes_docx(content: str, output_path: str):
    """
    将Markdown纪要内容生成专业Word文档
    content: 纪要Markdown文本
    output_path: 输出路径，如 /mnt/user-data/outputs/纪要.docx
    """
    doc = Document()
    
    # 页面设置：A4，页边距2cm
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    for margin in ['left_margin', 'right_margin', 'top_margin', 'bottom_margin']:
        setattr(section, margin, Cm(2.5))
    
    # 默认字体
    doc.styles['Normal'].font.name = '微软雅黑'
    doc.styles['Normal'].font.size = Pt(10.5)
    
    lines = content.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # 一级标题（文档标题）
        if line.startswith('# ') and not line.startswith('## '):
            p = doc.add_heading(line[2:], level=1)
            p.runs[0].font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            p.runs[0].font.size = Pt(16)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 元信息行（**时间：** 等）
        elif line.startswith('**') and '：' in line:
            p = doc.add_paragraph()
            # 解析加粗标签
            parts = re.split(r'\*\*(.*?)\*\*', line)
            for j, part in enumerate(parts):
                run = p.add_run(part)
                if j % 2 == 1:  # 奇数位是加粗内容
                    run.bold = True
                    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
                run.font.size = Pt(10.5)
        
        # 分隔线
        elif line == '---':
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(0)
        
        # 二级标题（## 一、核心摘要）
        elif line.startswith('## '):
            p = doc.add_heading(line[3:], level=2)
            p.runs[0].font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            p.runs[0].font.size = Pt(13)
        
        # 三级标题（### （一）业务板块）
        elif line.startswith('### '):
            p = doc.add_heading(line[4:], level=3)
            p.runs[0].font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
            p.runs[0].font.size = Pt(11)
        
        # 摘要条目（▶ 开头）
        elif line.startswith('▶ '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.space_before = Pt(4)
            content_line = line[2:]
            # 处理加粗
            parts = re.split(r'\*\*(.*?)\*\*', content_line)
            for j, part in enumerate(parts):
                run = p.add_run(part)
                if j % 2 == 1:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
                run.font.size = Pt(10.5)
        
        # 问题行（**Q：开头）
        elif line.startswith('**Q：') or line.startswith('**Q:'):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            # Q标签红色
            run_q = p.add_run('Q：')
            run_q.bold = True
            run_q.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            run_q.font.size = Pt(10.5)
            # 问题内容
            q_text = line.replace('**Q：', '').replace('**Q:', '').strip('*').strip()
            run_text = p.add_run(q_text)
            run_text.bold = True
            run_text.font.size = Pt(10.5)
        
        # 回答行（A：开头）
        elif line.startswith('A：') or line.startswith('A:'):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            # A标签蓝色
            run_a = p.add_run('A：')
            run_a.bold = True
            run_a.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            run_a.font.size = Pt(10.5)
            # 回答内容（处理加粗）
            a_text = line[2:].strip()
            parts = re.split(r'\*\*(.*?)\*\*', a_text)
            for j, part in enumerate(parts):
                run = p.add_run(part)
                if j % 2 == 1:
                    run.bold = True
                run.font.size = Pt(10.5)
        
        # 普通段落（含加粗处理）
        elif line:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            parts = re.split(r'\*\*(.*?)\*\*', line)
            for j, part in enumerate(parts):
                run = p.add_run(part)
                if j % 2 == 1:
                    run.bold = True
                run.font.size = Pt(10.5)
        
        i += 1
    
    doc.save(output_path)
    print(f"文档已保存至: {output_path}")
    return output_path


# 使用示例
if __name__ == '__main__':
    # 从文件读取纪要内容
    with open('/tmp/meeting_notes.md', 'r', encoding='utf-8') as f:
        content = f.read()
    
    output = create_meeting_notes_docx(
        content=content,
        output_path='/mnt/user-data/outputs/调研纪要.docx'
    )
```

## 执行步骤

1. 将Markdown纪要内容写入 `/tmp/meeting_notes.md`
2. 运行上述脚本
3. 使用 `present_files` 工具将生成的.docx提供给用户

## 文档样式说明

- **标题**：微软雅黑，居中，深蓝色（#1F497D）
- **一级标题**：16pt，居中
- **二级标题**：13pt，深蓝色
- **三级标题**：11pt，中蓝色（#2E74B5）
- **Q标签**：红色加粗（#C00000）
- **A标签**：深蓝色加粗（#1F497D）
- **摘要条目**：左缩进0.5cm
- **正文**：10.5pt微软雅黑
