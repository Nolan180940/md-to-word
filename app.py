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

# --- 3. 核心功能：智能修复引擎 (V7.1 列表/引用专项修复版) ---
def smart_fix_markdown(text):
    """
    使用逐行扫描 + 状态检测的方式修复 Markdown。
    重点修复：
    1. 列表变横杠问题 (通过强制前置空行修复)
    2. 引用块失效问题 (通过强制前置空行修复)
    3. 粗体/标题等格式粘连问题
    """
    if not text: return text, []
    
    log = []
    
    # 1. 全局清理：隐形字符
    if '\u200b' in text:
        text = text.replace('\u200b', '')
        log.append("🧹 移除了隐形字符")

    # 2. 全局清理：标准化 LaTeX 公式 (大模型常用方言)
    if '\\[' in text or '\\]' in text:
        text = text.replace('\\[', '$$').replace('\\]', '$$')
        log.append("📐 标准化块级公式")
    if '\\(' in text or '\\)' in text:
        text = text.replace('\\(', '$').replace('\\)', '$')
        log.append("📐 标准化行内公式")

    lines = text.split('\n')
    new_lines = []
    in_code_block = False  # 状态标记：是否在代码块内
    
    # 正则预编译
    re_code_fence = re.compile(r'^\s*```')
    re_heading = re.compile(r'^(#{1,6})([^ #])')     # 标题缺空格 #Title
    re_heading_std = re.compile(r'^(#{1,6}) (.*)')    # 标准标题
    
    # 引用正则：支持 >Text 和 > Text
    re_quote = re.compile(r'^(>+)([^ \n])')           # 引用缺空格 >Text
    re_quote_std = re.compile(r'^(>+)( .*)?')         # 标准引用 > Text
    
    # 列表正则：支持 -Item 和 - Item
    re_ul = re.compile(r'^(\s*[-*+])([^ \n])')        # 无序列表缺空格 -Item
    re_ul_std = re.compile(r'^(\s*[-*+]) (.*)')       # 标准无序列表 - Item
    
    re_ol = re.compile(r'^(\s*\d+\.)([^ \n])')        # 有序列表缺空格 1.Item
    re_ol_std = re.compile(r'^(\s*\d+\.) (.*)')       # 标准有序列表 1. Item
    
    re_hr = re.compile(r'^\s*([-*_]){3,}\s*$')        # 分割线
    re_bold_fix = re.compile(r'\*\*\s+(.*?)\s+\*\*')  # 修复粗体空格 ** text **

    for i, line in enumerate(lines):
        # --- A. 状态检测 ---
        # 如果遇到代码块标记，切换状态
        if re_code_fence.match(line):
            in_code_block = not in_code_block
            new_lines.append(line)
            continue
            
        # 如果在代码块内，直接保留原样，不做任何修改！
        if in_code_block:
            new_lines.append(line)
            continue

        # --- B. 行内格式修复 (仅在非代码块区域进行) ---
        
        # 1. 修复标题缺空格: #Title -> # Title
        if re_heading.match(line):
            line = re_heading.sub(r'\1 \2', line)
            if i < 5: log.append("🔨 修复了标题缺少空格")

        # 2. 修复引用缺空格: >Text -> > Text
        if re_quote.match(line):
            line = re_quote.sub(r'\1 \2', line)
            
        # 3. 修复列表缺空格: -Item -> - Item
        if re_ul.match(line):
            line = re_ul.sub(r'\1 \2', line)
        if re_ol.match(line):
            line = re_ol.sub(r'\1 \2', line)

        # 4. 修复粗体多余空格: ** text ** -> **text**
        # 很多时候粗体失效是因为这里多了空格
        if '**' in line:
            if re_bold_fix.search(line):
                line = re_bold_fix.sub(r'**\1**', line)

        # 5. 修复行内公式空格: $x$ -> $x$
        if '$' in line:
            line = re.sub(r'(?<!\$)\$[ \t]+(.*?)[ \t]+\$(?!\$)', r'$\1$', line)

        # 6. HTML 上标清理
        if '<sup>' in line:
            line = re.sub(r'<sup>(.*?)</sup>', r'^\1^', line)

        # --- C. 上下文空行注入 (解决粘连导致格式失效的核心逻辑) ---
        
        # 获取上一行内容 (如果存在)
        prev_line = lines[i-1] if i > 0 else ""
        is_prev_empty = not prev_line.strip()
        
        # 规则1: 引用块隔离
        # 逻辑：如果当前是引用，且上一行不是引用、不是空行 -> 加空行
        # 这确保了 Pandoc 能识别出这是一个新的 Blockquote 块
        if re_quote_std.match(line):
            if not is_prev_empty and not re_quote_std.match(prev_line):
                new_lines.append("") 
        
        # 规则2: 列表隔离 (关键修复：让横杠变成圆点)
        # 逻辑：如果当前是列表，且上一行不是同类型的列表、不是空行 -> 加空行
        # Pandoc 要求列表前必须有空行，否则会被当作普通文本处理
        elif re_ul_std.match(line):
            is_prev_ul = re_ul_std.match(prev_line)
            if not is_prev_empty and not is_prev_ul:
                new_lines.append("")
        elif re_ol_std.match(line):
            is_prev_ol = re_ol_std.match(prev_line)
            if not is_prev_empty and not is_prev_ol:
                new_lines.append("")

        # 规则3: 标题隔离
        # 标题前面必须有空行
        elif re_heading_std.match(line):
            if not is_prev_empty:
                new_lines.append("")

        # 规则4: 分割线隔离
        # 分割线前面必须有空行
        elif re_hr.match(line):
            if not is_prev_empty:
                new_lines.append("")
            
        new_lines.append(line)
        
        # 规则5: 分割线后也强制加空行
        if re_hr.match(line):
            new_lines.append("")

    # 4. 重新组合
    fixed_text = "\n".join(new_lines)
    
    # 5. 收尾：代码块闭合检查
    # 如果代码块状态最后还是 True，说明漏了闭合
    if in_code_block:
        fixed_text += "\n```"
        log.append("🧱 自动闭合了未结束的代码块")

    # 6. 大扫除：清理超过3个的连续换行，保持整洁
    fixed_text = re.sub(r'\n{4,}', r'\n\n', fixed_text)

    return fixed_text, list(set(log))

# --- 4. 核心功能：Word 样式后处理 (完全保留原版) ---
def apply_word_styles(docx_path):
    if not HAS_DOCX:
        return
        
    try:
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
    except Exception as e:
        print(f"Docx处理错误: {e}")

# --- 5. 转换与生成 (完全保留原版) ---
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

st.title("🛠️ Markdown 转 Word 稳定版")
st.caption("代码块阴影 | 引用块缩进(正体) | 智能修复标题/列表/引用/分割线")
st.divider()

if not HAS_DOCX:
    st.error("⚠️ 检测到未安装 `python-docx` 库。样式增强功能将无法生效。")

# 默认示例文本
default_text = r'''# 格式修复测试

## 1. 粗体修复
这里的** 粗体 **中间有多余空格，以前会挂，现在应该能自动修复为**粗体**。

## 2. 列表修复 (粘连测试)
上一行是文本，下一行直接开始列表(没有空行)：
- 这是列表项1
- 这是列表项2
(现在应该能自动在上面插入空行，变成真正的圆点列表)

## 3. 引用修复
> 这是引用块
>
> 这是第二行
(现在应该能保持连贯，且有灰色缩进样式)

## 4. 代码块
```python
print("Hello World")
```
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
