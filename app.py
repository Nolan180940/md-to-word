import streamlit as st
import pypandoc
import tempfile
import os
import re

# 尝试导入 python-docx，用于后期处理 Word 样式
try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# --- 1. 页面配置 ---
st.set_page_config(
    page_title="Markdown to Word Pro (智能修复版)",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. CSS 美化 ---
st.markdown("""
<style>
    h1, h2, h3 { font-family: 'Segoe UI', sans-serif; font-weight: 600; }
    .stTextArea textarea { font-family: 'Consolas', monospace; font-size: 14px; }
    .fix-report {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 8px;
        padding: 10px;
        color: #166534;
        font-size: 0.9em;
        margin-bottom: 10px;
    }
    .fix-report-item { margin-left: 1em; }
    @media (prefers-color-scheme: dark) {
        .fix-report { background-color: #064e3b; border-color: #065f46; color: #ecfccb; }
    }
</style>
""", unsafe_allow_html=True)

# --- 3. 核心功能：智能修复引擎 (V5.1 增强版) ---
def smart_fix_markdown(text):
    log = []
    fixed_text = text

    # 1. [基础] 清理零宽空格
    if '\u200b' in fixed_text:
        fixed_text = fixed_text.replace('\u200b', '')
        log.append("🧹 移除了隐形字符")

    # 2. [关键] 强制标准化 LaTeX 公式语法
    # 处理块级公式 \[ ... \] -> $$...$$
    if '\\[' in fixed_text or '\\]' in fixed_text:
        fixed_text = fixed_text.replace('\\[', '$$').replace('\\]', '$$')
        log.append("📐 将 LaTeX 块级公式 \\[...\\] 标准化为 $$...$$")

    # 处理行内公式 \( ... \) -> $...$
    if '\\(' in fixed_text or '\\)' in fixed_text:
        fixed_text = fixed_text.replace('\\(', '$').replace('\\)', '$')
        log.append("📐 将 LaTeX 行内公式 \\(...\\) 标准化为 $...$")

    # 3. [新增] 修复行内公式多余空格 $x$ -> $x$
    # Pandoc 对 inline math 要求 $ 后紧跟内容，$ 前紧跟内容
    # 正则说明：(?<!\$) 排除 $$ 的情况
    # \$[ \t]+ 匹配起始 $ 后的空格
    # (.*?) 捕获内容
    # [ \t]+\$ 匹配结束 $ 前的空格
    pattern_space_math = r'(?<!\$)\$[ \t]+(.*?)[ \t]+\$(?!\$)'
    if re.search(pattern_space_math, fixed_text):
        new_text, count = re.subn(pattern_space_math, r'$\1$', fixed_text)
        if count > 0:
            fixed_text = new_text
            log.append(f"🔧 移除了 {count} 处行内公式的多余空格 ($x$ -> $x$)")

    # 4. [HTML 清理] 将 <sup>...</sup> 转换为 Pandoc 上标 ^...^
    if '<sup>' in fixed_text:
        new_text, count = re.subn(r'<sup>(.*?)</sup>', r'^\1^', fixed_text)
        if count > 0:
            fixed_text = new_text
            log.append(f"⬆️ 将 {count} 处 HTML 上标标签转换为 Markdown 格式")

    # 5. [闭合检查] 自动闭合代码块
    code_fence_count = len(re.findall(r'^```', fixed_text, re.MULTILINE))
    if code_fence_count % 2 != 0:
        fixed_text += "\n```"
        log.append("🧱 自动闭合了未结束的代码块")

    # 6. [闭合检查] 自动闭合公式块
    math_block_count = fixed_text.count('$$')
    if math_block_count % 2 != 0:
        fixed_text += "\n$$"
        log.append("🧮 自动闭合了未结束的 LaTeX 公式块")
        
    # 7. [格式优化] 确保代码块前后有空行
    fixed_text = re.sub(r'([^\n])\n```', r'\1\n\n```', fixed_text)
    fixed_text = re.sub(r'```\n([^\n])', r'```\n\n\1', fixed_text)
    
    # 8. [新增] 确保引用块前后有空行
    # 匹配以 > 开头的行（可能前面有空格），并在其前后添加空行
    # 处理引用块前的空行
    fixed_text = re.sub(r'([^\n>])\n[ \t]*>', r'\1\n\n>', fixed_text, flags=re.MULTILINE)
    # 处理引用块后的空行（在非引用行前添加空行）
    fixed_text = re.sub(r'>[^\n]*\n([^\n>])', r'>\g<0>\n\1', fixed_text, flags=re.MULTILINE)
    # 为连续的引用行组后添加空行
    fixed_text = re.sub(r'([>].*?)(\n(?![> \t]))', r'\1\2\n', fixed_text, flags=re.MULTILINE)
    # 处理引用块后紧跟非引用行的情况
    fixed_text = re.sub(r'([>][^\n]*\n)([^\n> \t])', r'\1\n\2', fixed_text, flags=re.MULTILINE)
    
    # 更精确的引用块处理：查找整个引用块并确保其前后有空行
    # 首先，确保引用块之前有空行（如果前面不是空行或另一个引用）
    fixed_text = re.sub(r'([^\n> \t])\n([ \t]*>[^\n]*(?:\n[ \t]*>[^\n]*)*)', r'\1\n\n\2', fixed_text, flags=re.MULTILINE)
    # 然后，确保引用块之后有空行（如果后面不是空行或另一个引用）
    # 使用更复杂的正则来匹配完整的引用块
    original_text = fixed_text
    # 处理引用块后的情况
    lines = fixed_text.split('\n')
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # 检查是否是引用行
        if line.strip().startswith('>'):
            # 收集连续的引用行
            block_lines = []
            j = i
            while j < len(lines) and (lines[j].strip().startswith('>') or (lines[j].strip() == '' and j+1 < len(lines) and lines[j+1].strip().startswith('>'))):
                block_lines.append(lines[j])
                j += 1
            
            # 检查当前块前是否有内容且不是空行
            if i > 0 and lines[i-1].strip() != '':
                # 在块前插入空行
                if new_lines and new_lines[-1] != '':
                    new_lines.append('')
            
            # 添加引用块
            new_lines.extend(block_lines)
            
            # 检查块后是否有内容且不是空行也不是另一个引用块
            if j < len(lines) and lines[j].strip() != '' and not lines[j].strip().startswith('>'):
                # 在块后插入空行
                new_lines.append('')
            
            i = j
        else:
            new_lines.append(line)
            i += 1
    
    fixed_text = '\n'.join(new_lines)
    
    # 记录引用块修复日志
    quote_block_count = len(re.findall(r'^[ \t]*>[^\n]*', original_text, re.MULTILINE))
    if quote_block_count > 0:
        log.append(f"💬 处理了 {quote_block_count} 个引用块，确保其前后有空行")

    return fixed_text, log

# --- 4. 核心功能：Word 样式后处理 ---
def apply_word_styles(docx_path):
    if not HAS_DOCX:
        return 
        
    doc = Document(docx_path)
    styles = doc.styles

    # === 1. 优化代码块样式 (Source Code) ===
    try:
        style_name = 'Source Code' if 'Source Code' in styles else 'SourceCode'
        if style_name in styles:
            style_code = styles[style_name]
            style_code.font.name = 'Consolas'
            style_code.font.size = Pt(10)
            
            p_pr = style_code.element.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'F2F2F2') 
            p_pr.append(shd)
            
            if not p_pr.find(qn('w:pBdr')):
                pbdr = OxmlElement('w:pBdr')
                for border in ['top', 'left', 'bottom', 'right']:
                    b = OxmlElement(f'w:{border}')
                    b.set(qn('w:val'), 'single')
                    b.set(qn('w:sz'), '4') 
                    b.set(qn('w:space'), '1')
                    b.set(qn('w:color'), 'D4D4D4') 
                    pbdr.append(b)
                p_pr.append(pbdr)
            
    except Exception as e:
        print(f"代码块样式应用失败: {e}")

    # === 2. 优化引用块样式 (Block Text) ===
    try:
        target_styles = ['Block Text', 'Quote', 'BlockText']
        found_style = None
        for name in target_styles:
            if name in styles:
                found_style = styles[name]
                break
        
        if found_style:
            found_style.font.color.rgb = RGBColor(105, 105, 105) 
            found_style.font.italic = False
            found_style.paragraph_format.left_indent = Inches(0.25)
            
            p_pr = found_style.element.get_or_add_pPr()
            if not p_pr.find(qn('w:pBdr')):
                pbdr = OxmlElement('w:pBdr')
                left = OxmlElement('w:left')
                left.set(qn('w:val'), 'single')
                left.set(qn('w:sz'), '12') 
                left.set(qn('w:space'), '12') 
                left.set(qn('w:color'), '999999') 
                pbdr.append(left)
                p_pr.append(pbdr)

    except Exception as e:
        print(f"引用样式应用失败: {e}")

    doc.save(docx_path)

# --- 5. 转换与生成 ---
def convert_to_docx(md_content):
    output_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            output_path = tmp_file.name
        
        pypandoc.convert_text(
            md_content, 
            'docx', 
            format='markdown+tex_math_dollars', 
            outputfile=output_path, 
            extra_args=['--standalone']
        )
        
        if HAS_DOCX:
            apply_word_styles(output_path)
            
        return output_path, None
    except Exception as e:
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
        return None, str(e)

# --- 6. 智能文件名生成 ---
def generate_smart_filename(text):
    if not text or not text.strip():
        return "document.docx"
    
    h1_match = re.search(r'^#\s+(.+)$', text, re.MULTILINE)
    if h1_match:
        raw_title = h1_match.group(1).strip()
    else:
        h2_match = re.search(r'^##\s+(.+)$', text, re.MULTILINE)
        if h2_match:
            raw_title = h2_match.group(1).strip()
        else:
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            raw_title = lines[0] if lines else "document"

    clean_name = re.sub(r'[\\/*?:"<>|]', '', raw_title)
    clean_name = re.sub(r'[*_`]', '', clean_name)
    final_name = clean_name[:40].strip()
    
    return f"{final_name}.docx"

# --- 7. 界面布局 ---

st.title("🛠️ Markdown 转 Word")
st.caption("代码块阴影 | 引用块缩进(正体) | 智能标题生成 | 自动修复公式空格")
st.divider()

if not HAS_DOCX:
    st.error("⚠️ 检测到未安装 `python-docx` 库。样式增强功能将无法生效。")

# 默认示例文本
default_text = r'''# 深度学习中的概率分布

这是一个包含 "空格公式" 的测试。

## 1. 坏掉的公式 (Spaces)

大模型经常输出这种带空格的行内公式： $E = mc^2$ ，或者 $ x_0 = 0 $。
在 Pandoc 里，这通常会被解析成普通文本。

本工具会自动将其修复为：$E=mc^2$ 和 $x_0=0$。

## 2. 块级公式 (LaTeX 风格)

\[
\mathcal{L}(\theta) = -\frac{1}{N} \sum_{i=1}^N \left[ y_i \log(\hat{y}_i) + (1-y_i) \log(1-\hat{y}_i) \right]
\]

## 3. 代码块测试

```python
def fix_spaces(text):
    return text.strip()
