import streamlit as st
import pypandoc
import tempfile
import os
import re

# 尝试导入 python-docx
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
    page_title="Markdown to Word Pro (稳定版)",
    page_icon="🛡️",
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

# --- 3. 核心功能：智能修复引擎 (V6.4 强力空行版) ---
def smart_fix_markdown(text):
    log = []
    fixed_text = text

    # 1. [基础] 清理零宽空格
    if '\u200b' in fixed_text:
        fixed_text = fixed_text.replace('\u200b', '')
        log.append("🧹 移除了隐形字符")

    # 2. [关键] 强制修复标题语法 (#Title / #   Title -> # Title)
    # 逻辑：查找行首的 #，无论后面有没有空格，或者有多少空格，统一规范化为 "# "
    # 同时确保标题前面有空行 (除非是文件第一行)
    
    # 2.1 修复缺少空格 (#Title -> # Title)
    pattern_heading_missing = r'^(#{1,6})([^ \t\n#])'
    if re.search(pattern_heading_missing, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_heading_missing, r'\1 \2', fixed_text, flags=re.MULTILINE)
        log.append("🔨 修复了标题缺少空格的问题")
    
    # 2.2 修复多余空格 (#   Title -> # Title)
    pattern_heading_extra = r'^(#{1,6})[ \t]{2,}'
    if re.search(pattern_heading_extra, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_heading_extra, r'\1 ', fixed_text, flags=re.MULTILINE)
        log.append("🔨 标准化了标题空格")

    # 2.3 [新增] 强制标题前加空行 (防止粘连上一段)
    # 查找：非换行符 + 换行 + #号
    fixed_text = re.sub(r'([^\n])\n(#{1,6}[ \t])', r'\1\n\n\2', fixed_text)

    # 3. [关键] 强制修复引用语法 (>Text -> > Text)
    # 3.1 修复缺少空格
    pattern_quote = r'^(>+)([^ \t\n])'
    if re.search(pattern_quote, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_quote, r'\1 \2', fixed_text, flags=re.MULTILINE)
        log.append("🔨 修复了引用缺少空格的问题")

    # 3.2 [新增] 强制引用块前加空行
    # 查找：非换行符 + 换行 + >号
    fixed_text = re.sub(r'([^\n])\n(>+[ \t])', r'\1\n\n\2', fixed_text)

    # 4. [列表] 修复列表语法 (-Item -> - Item)
    pattern_ul = r'^(\s*[-*+])([^ \t\n])'
    if re.search(pattern_ul, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_ul, r'\1 \2', fixed_text, flags=re.MULTILINE)
        log.append("📋 修复了粘连的无序列表语法")
    
    pattern_ol = r'^(\s*\d+\.)([^ \t\n])'
    if re.search(pattern_ol, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_ol, r'\1 \2', fixed_text, flags=re.MULTILINE)
        log.append("🔢 修复了粘连的有序列表语法")

    # 5. [关键] 强制修复分割线 (---)
    # 强制在分割线前后各加两个换行符
    pattern_hr = r'^\s*([-*_]){3,}\s*$'
    if re.search(pattern_hr, fixed_text, re.MULTILINE):
        fixed_text = re.sub(pattern_hr, r'\n\n---\n\n', fixed_text, flags=re.MULTILINE)
        log.append("➖ 强制分割线前后换行")

    # 6. [LaTeX] 强制标准化公式语法
    if '\\[' in fixed_text or '\\]' in fixed_text:
        fixed_text = fixed_text.replace('\\[', '$$').replace('\\]', '$$')
        log.append("📐 标准化块级公式")
    if '\\(' in fixed_text or '\\)' in fixed_text:
        fixed_text = fixed_text.replace('\\(', '$').replace('\\)', '$')
        log.append("📐 标准化行内公式")

    # 7. [LaTeX] 修复行内公式多余空格
    pattern_space_math = r'(?<!\$)\$[ \t]+(.*?)[ \t]+\$(?!\$)'
    if re.search(pattern_space_math, fixed_text):
        fixed_text = re.sub(pattern_space_math, r'$\1$', fixed_text)
        log.append("🔧 移除了行内公式的多余空格")

    # 8. [HTML] 清理上标
    if '<sup>' in fixed_text:
        fixed_text = re.sub(r'<sup>(.*?)</sup>', r'^\1^', fixed_text)
        log.append("⬆️ 转换 HTML 上标")

    # 9. [闭合] 自动闭合代码块/公式
    code_fence_count = len(re.findall(r'^```', fixed_text, re.MULTILINE))
    if code_fence_count % 2 != 0:
        fixed_text += "\n```"
        log.append("🧱 自动闭合代码块")
    
    # 10. [格式] 代码块前后强制空行
    fixed_text = re.sub(r'([^\n])\n```', r'\1\n\n```', fixed_text)
    fixed_text = re.sub(r'```\n([^\n])', r'```\n\n\1', fixed_text)
    
    # 11. [大扫除] 清理过多空行
    fixed_text = re.sub(r'\n{4,}', r'\n\n', fixed_text)
    
    return fixed_text, log

# --- 4. 核心功能：Word 样式后处理 (保持原样) ---
def apply_word_styles(docx_path):
    if not HAS_DOCX:
        return
        
    try:
        doc = Document(docx_path)
        styles = doc.styles

        # === 1. 优化代码块样式 ===
        target_code_styles = ['Source Code', 'SourceCode', 'Verbatim Char']
        
        for name in target_code_styles:
            if name in styles:
                try:
                    found_code_style = styles[name]
                    found_code_style.font.name = 'Consolas'
                    found_code_style.font.size = Pt(10)
                    
                    p_pr = found_code_style.element.get_or_add_pPr()
                    
                    if not p_pr.find(qn('w:shd')):
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
                except Exception:
                    continue 

        # === 2. 优化引用块样式 ===
        target_quote_styles = ['Block Text', 'Quote', 'BlockText']
        
        for name in target_quote_styles:
            if name in styles:
                try:
                    found_style = styles[name]
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
                except Exception:
                    continue

        doc.save(docx_path)
    except Exception as e:
        print(f"样式应用失败 (非致命错误): {e}")

# --- 5. 转换与生成 (保持原样) ---
def convert_to_docx(md_content):
    output_path = None
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            output_path = tmp_file.name
        
        # 1. 基础转换 (Pandoc)
        pypandoc.convert_text(
            md_content, 
            'docx', 
            format='markdown+tex_math_dollars+tex_math_single_backslash', 
            outputfile=output_path, 
            extra_args=['--standalone']
        )
        
        # 2. 样式增强
        if HAS_DOCX:
            try:
                apply_word_styles(output_path)
            except Exception as e:
                print(f"样式增强失败，降级为普通 Word: {e}")
            
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

st.title("🛠️ Markdown 转 Word 稳定版")
st.caption("代码块阴影 | 引用块缩进(正体) | 智能修复标题/列表/引用/分割线")
st.divider()

if not HAS_DOCX:
    st.error("⚠️ 检测到未安装 `python-docx` 库。样式增强功能将无法生效。")

# 默认示例文本
default_text = r'''# 格式大乱斗测试

##标题粘连测试(应该自动修复)
这里没有空格，普通Markdown解析器会挂。

>引用粘连测试(应该自动修复)
>也没有空格。

-无序列表粘连测试
1.有序列表粘连测试

---
上面是粘连的分割线(应该自动变成横线)。

```python
def code():
    pass
# 后面少写了闭合
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

# --- 底部 ---
st.divider()
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    if st.button("🚀 生成 Word 文档", type="primary", use_container_width=True):
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
