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
    if '\\[' in fixed_text or '\\]' in fixed_text:
        fixed_text = fixed_text.replace('\\[', '$$').replace('\\]', '$$')
        log.append("📐 将 LaTeX 块级公式 \\[...\\] 标准化为 $$...$$")

    if '\\(' in fixed_text or '\\)' in fixed_text:
        fixed_text = fixed_text.replace('\\(', '$').replace('\\)', '$')
        log.append("📐 将 LaTeX 行内公式 \\(...\\) 标准化为 $...$")

    # 3. [新增] 修复行内公式多余空格
    pattern_space_math = r'(?<!\$)\$[ \t]+(.*?)[ \t]+\$(?!\$)'
    new_text, count = re.subn(pattern_space_math, r'$\1$', fixed_text)
    if count > 0:
        fixed_text = new_text
        log.append(f"🔧 移除了 {count} 处行内公式的多余空格")

    # 4. [HTML 清理] 修复上标
    new_text, count = re.subn(r'<sup>(.*?)</sup>', r'^\1^', fixed_text)
    if count > 0:
        fixed_text = new_text
        log.append(f"⬆️ 将 {count} 处 HTML 上标转换为 Markdown")

    # 5. 自动闭合代码块
    if len(re.findall(r'^```', fixed_text, re.MULTILINE)) % 2 != 0:
        fixed_text += "\n```"
        log.append("🧱 自动闭合了未结束的代码块")

    # 6. 自动闭合公式块
    if fixed_text.count('$$') % 2 != 0:
        fixed_text += "\n$$"
        log.append("🧮 自动闭合了未结束的公式块")

    # 7. 确保代码块前后有空行
    fixed_text = re.sub(r'([^\n])\n```', r'\1\n\n```', fixed_text)
    fixed_text = re.sub(r'```\n([^\n])', r'```\n\n\1', fixed_text)

    # 8. **关键新增：blockquote 段落前后加空行**
    blockquote_pattern = r'(?:^>.*(?:\n|$))+'
    matches = list(re.finditer(blockquote_pattern, fixed_text, re.MULTILINE))

    for m in reversed(matches):
        start, end = m.start(), m.end()
        before = fixed_text[:start].rstrip('\n')
        quote_block = fixed_text[start:end].rstrip('\n')
        after = fixed_text[end:].lstrip('\n')

        if not before.endswith('\n\n'):
            before += "\n\n"
            log.append("🧩 在 blockquote 前加入空行")

        if after and not after.startswith('\n\n'):
            after = "\n\n" + after
            log.append("🧩 在 blockquote 后加入空行")

        fixed_text = before + "\n" + quote_block + "\n" + after

    return fixed_text, log

# --- 4. Word 样式后处理 ---
def apply_word_styles(docx_path):
    if not HAS_DOCX:
        return

    doc = Document(docx_path)
    styles = doc.styles

    # 代码块样式优化
    try:
        for s in ['Source Code', 'SourceCode', 'Source Code Char']:
            if s in styles:
                style = styles[s]
                style.font.name = 'Consolas'
                style.font.size = Pt(10)

                p_pr = style.element.get_or_add_pPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:fill'), 'F2F2F2')
                p_pr.append(shd)
                break
    except:
        pass

    # Quote/引用块样式优化
    try:
        for s in ['Block Text', 'Quote', 'BlockText']:
            if s in styles:
                style = styles[s]
                style.font.italic = False
                style.font.color.rgb = RGBColor(105, 105, 105)
                style.paragraph_format.left_indent = Inches(0.25)
                break
    except:
        pass

    doc.save(docx_path)

# --- 5. Pandoc 生成 docx ---
def convert_to_docx(md_content):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = tmp.name

        pypandoc.convert_text(
            md_content,
            'docx',
            format='markdown+tex_math_dollars',
            outputfile=output_path,
            extra_args=['--standalone']
        )

        apply_word_styles(output_path)
        return output_path, None
    except Exception as e:
        return None, str(e)

# --- 6. 智能文件名生成（可靠版） ---
def generate_smart_filename(text):
    if not text or not text.strip():
        return "document.docx"

    h1 = re.search(r'^\s*#\s+(.+)$', text, re.MULTILINE)
    if h1:
        title = h1.group(1).strip()
    else:
        h2 = re.search(r'^\s*##\s+(.+)$', text, re.MULTILINE)
        if h2:
            title = h2.group(1).strip()
        else:
            title = next((l.strip() for l in text.splitlines() if l.strip()), "document")

    title = re.sub(r'[\\/*?:"<>|]', '', title)
    title = re.sub(r'[*_`]', '', title)
    title = title[:40].strip()
    if not title:
        title = "document"

    return f"{title}.docx"

# --- 7. UI 界面 ---
st.title("🛠️ Markdown 转 Word")
st.caption("代码块阴影 | 引用块缩进(正体) | 智能标题生成 | 自动修复公式空格")
st.divider()

default_text = r'''# 示例标题
这里是内容

> 这是一个引用测试
> 这里是连续多行 blockquote

下文内容应该与引用可靠分隔。

'''

col_input, col_preview = st.columns(2, gap="medium")

with col_input:
    st.subheader("⌨️ 编辑区")
    md_text = st.text_area("Input", value=default_text, height=600, label_visibility="collapsed")

with col_preview:
    st.subheader("👁️ 预览 (修复后)")
    preview_text, logs = smart_fix_markdown(md_text)

    if logs:
        with st.expander(f"🤖 自动执行了 {len(logs)} 项智能修复", expanded=True):
            for item in logs:
                st.markdown(f"- {item}")

    with st.container(border=True):
        if preview_text.strip():
            st.markdown(preview_text)
        else:
            st.write("等待输入...")

st.divider()

with st.columns([1,2,1])[1]:
    if st.button("🚀 生成定制化 Word 文档", type="primary", use_container_width=True):
        if not md_text.strip():
            st.warning("⚠️ 内容不能为空")
        else:
            final_text, _ = smart_fix_markdown(md_text)
            file_name = generate_smart_filename(final_text)

            docx_path, error_msg = convert_to_docx(final_text)

            if docx_path:
                with open(docx_path, "rb") as f:
                    data = f.read()

                st.success(f"✅ 生成成功：{file_name}")
                st.download_button("⬇️ 下载 Word", data=data, file_name=file_name, mime="application/docx")

                try:
                    os.remove(docx_path)
                except:
                    pass
            else:
                st.error("❌ 转换失败")
                if error_msg:
                    st.code(error_msg)
