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

# --- 页面配置 ---
st.set_page_config(
    page_title="Markdown to Word Pro (智能修复版)",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

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

# =============================
# smart_fix_markdown（改进版）
# - 更保守地处理行内/块级公式
# - 将 \[...\] 转为 $$\n...\n$$ 并保留换行
# - 只在必要情况下修复行内公式的多余空格（不跨行）
# - 确保连续以 '>' 开头的 blockquote 段落在前后各有空行（不改变 '>'）
# - 自动闭合未配对的 ``` 和 $$（对 $$ 以双 $ 为单位）
# =============================
def smart_fix_markdown(text):
    log = []
    fixed_text = text if text is not None else ""

    # 1) 清理零宽空格（不可见字符）
    if '\u200b' in fixed_text:
        fixed_text = fixed_text.replace('\u200b', '')
        log.append("🧹 移除了隐形字符")

    # 2) 将 \[ ... \]（LaTeX 块）转为 $$\n ... \n$$ （保留内部换行）
    # 使用 DOTALL 捕获跨行内容
    def repl_bracket_block(m):
        inner = m.group(1).rstrip('\n')
        return "$$\n" + inner + "\n$$"
    new_text, cnt = re.subn(r'\\\[(.*?)\\\]', repl_bracket_block, fixed_text, flags=re.DOTALL)
    if cnt > 0:
        fixed_text = new_text
        log.append(f"📐 将 {cnt} 处 LaTeX 块 \\[...\\] 标准化为 $$...$$（并保留换行）")

    # 3) 将 \( ... \)（短 inline LaTeX）转为 $...$ —— 但只在不跨行时替换
    new_text, cnt = re.subn(r'\\\(([^\n]*?)\\\)', r'$\1$', fixed_text)
    if cnt > 0:
        fixed_text = new_text
        log.append(f"📐 将 {cnt} 处 LaTeX 行内公式 \\(...\\) 标准化为 $...$")

    # 4) 修复行内公式两端多余空格，但严格不跨行（避免把跨行文本吞掉）
    pattern_space_math = r'(?<!\$)\$[ \t]+([^\n]*?)[ \t]+\$(?!\$)'
    new_text, cnt = re.subn(pattern_space_math, r'$\1$', fixed_text)
    if cnt > 0:
        fixed_text = new_text
        log.append(f"🔧 移除了 {cnt} 处行内公式的多余空格（仅单行匹配）")

    # 5) HTML 上标 <sup>...</sup> -> ^...^ （宽松替换）
    new_text, cnt = re.subn(r'<sup>(.*?)</sup>', r'^\1^', fixed_text, flags=re.DOTALL)
    if cnt > 0:
        fixed_text = new_text
        log.append(f"⬆️ 将 {cnt} 处 HTML 上标转换为 Markdown 上标格式")

    # 6) 确保代码块 ``` 的数量为偶数（若为奇数则在文本末尾补一个 ```）
    code_fence_count = len(re.findall(r'^```', fixed_text, re.MULTILINE))
    if code_fence_count % 2 != 0:
        fixed_text = fixed_text.rstrip('\n') + "\n\n```\n"
        log.append("🧱 自动闭合了未结束的代码块（在文末补 ```）")

    # 7) 针对 $$ 的块公式配对（只统计完整的 $$ 符号对）
    #    先查找所有 $$ 的位置，若为奇数则在文末补一个 $$（不会尝试修复单个 $）
    dollar_pairs = re.findall(r'\$\$', fixed_text)
    if len(dollar_pairs) % 2 != 0:
        fixed_text = fixed_text.rstrip('\n') + "\n\n$$\n"
        log.append("🧮 自动闭合了未结束的 $$ 块公式（在文末补 $$）")

    # 8) 确保代码块前后有空行（增强 Pandoc 识别）
    fixed_text = re.sub(r'([^\n])\n```', r'\1\n\n```', fixed_text)
    fixed_text = re.sub(r'```\n([^\n])', r'```\n\n\1', fixed_text)

    # 9) 最保守的 blockquote 处理：按行扫描，合并连续 '>' 行为一个 block，
    #    在 block 前后各插入一个空行（如果本来就有空行则不重复插入）。
    lines = fixed_text.splitlines()
    if lines:
        out = []
        i = 0
        made_change = False
        while i < len(lines):
            line = lines[i]
            if line.lstrip().startswith('>'):
                # 前置空行
                if out and out[-1].strip() != '':
                    out.append('')
                    made_change = True

                # 把连续的 > 行直接复制到 out
                while i < len(lines) and lines[i].lstrip().startswith('>'):
                    out.append(lines[i])
                    i += 1

                # 后置空行（仅当后面还有非空内容时）
                if i < len(lines) and lines[i].strip() != '':
                    out.append('')
                    made_change = True
                continue
            else:
                out.append(line)
                i += 1

        # 恢复末尾换行状态
        if fixed_text.endswith('\n'):
            new_fixed_text = '\n'.join(out) + '\n'
        else:
            new_fixed_text = '\n'.join(out)

        if made_change:
            fixed_text = new_fixed_text
            log.append("🧩 在 blockquote 段落的前后强制加入空行（仅插入必要的空行）")

    # 返回修复后的文本与日志
    return fixed_text, log

# ========== 以下为原有代码（未变） ==========

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

# ========== UI 布局（不变） ==========
st.title("🛠️ Markdown 转 Word")
st.caption("代码块阴影 | 引用块缩进(正体) | 智能标题生成 | 自动修复公式空格")
st.divider()

if not HAS_DOCX:
    st.error("⚠️ 检测到未安装 `python-docx` 库。样式增强功能将无法生效。")

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
'''

col_input, col_preview = st.columns(2, gap="medium")

with col_input:
    st.subheader("⌨️ 编辑区")
    md_text = st.text_area(
        "Input", 
        value=default_text, 
        height=600, 
        label_visibility="collapsed",
        placeholder="在此粘贴..."
    )

with col_preview:
    st.subheader("👁️ 实时预览 (修复后)")
    preview_text, logs = smart_fix_markdown(md_text)

    if logs:
        with st.expander(f"🤖 自动执行了 {len(logs)} 项智能修复", expanded=True):
            for log in logs:
                st.markdown(f"- {log}")

    with st.container(border=True):
        if preview_text.strip():
            st.markdown(preview_text, unsafe_allow_html=True)
        else:
            st.write("等待输入...")

st.divider()
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    if st.button("🚀 生成定制化 Word 文档", type="primary", use_container_width=True):
        if not md_text.strip():
            st.warning("⚠️ 内容不能为空")
        else:
            final_text, _ = smart_fix_markdown(md_text)
            file_name = generate_smart_filename(final_text)

            with st.spinner("正在渲染并注入样式..."):
                docx_path, error_msg = convert_to_docx(final_text)
                
            if docx_path and os.path.exists(docx_path):
                with open(docx_path, "rb") as f:
                    file_data = f.read()
                
                st.success(f"✅ 生成成功！文件名为：**{file_name}**")
                
                st.download_button(
                    label="⬇️ 点击下载 Word",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                try:
                    os.remove(docx_path)
                except:
                    pass
            else:
                st.error("❌ 转换失败")
                if error_msg:
                    st.code(error_msg)
