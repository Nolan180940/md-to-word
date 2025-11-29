import streamlit as st
import pypandoc
import tempfile
import os

# --- 1. 页面配置 (强制宽屏 + 暗色兼容) ---
st.set_page_config(
    page_title="Markdown to Word Pro",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. 极简 CSS 美化 (仅用于字体和微调，不破坏布局) ---
st.markdown("""
<style>
    /* 调整一下标题的字体，更有科技感 */
    h1, h2, h3 {
        font-family: 'Segoe UI', sans-serif;
        font-weight: 600;
    }
    /* 让输入框的代码字体更好看 */
    .stTextArea textarea {
        font-family: 'Consolas', monospace;
    }
    /* 调整一下成功消息的样式 */
    .stAlert {
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. 核心功能：Pandoc 转换 ---
def convert_to_docx(md_content):
    """
    使用 Pandoc 将 Markdown 转换为 Docx
    """
    try:
        # 创建临时文件
        # delete=False 是为了兼容 Windows，Windows 下不能在文件打开时删除
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            output_path = tmp_file.name
        
        # 执行转换
        # extra_args=['--standalone'] 确保生成完整的文档结构
        pypandoc.convert_text(
            md_content, 
            'docx', 
            format='markdown', 
            outputfile=output_path, 
            extra_args=['--standalone']
        )
        return output_path
    except Exception as e:
        st.error(f"转换引擎出错: {str(e)}")
        st.info("💡 请确认您的电脑已安装 Pandoc (https://pandoc.org/installing.html)")
        return None

# --- 4. 界面布局 ---

# 标题区
st.title("📝 Markdown 转 Word ")
st.caption("所见即所得 | 完美支持 LaTeX 数学公式")
st.divider()

# 默认示例文本
default_text = r"""
# 🚀 欢迎使用

这是一段测试文本。您可以在左侧输入 Markdown，右侧会实时显示渲染结果。

## 1. 数学公式支持

著名的麦克斯韦方程组 (Maxwell's Equations):

$$
\begin{aligned}
\nabla \cdot \mathbf{E} &= \frac{\rho}{\varepsilon_0} \\
\nabla \cdot \mathbf{B} &= 0 \\
\nabla \times \mathbf{E} &= -\frac{\partial \mathbf{B}}{\partial t} \\
\nabla \times \mathbf{B} &= \mu_0\mathbf{J} + \mu_0\varepsilon_0\frac{\partial \mathbf{E}}{\partial t}
\end{aligned}
$$

以及行内公式：例如欧拉公式 $e^{i\pi} + 1 = 0$。

## 2. 代码高亮

```python
import numpy as np

def sigmoid(x):
    return 1 / (1 + np.exp(-x))
```

## 3. 列表与引用

- 支持无序列表
- 支持有序列表

> 这是一个引用块，转换到 Word 后会保持引用样式。
"""

# 主体布局：两列
# 使用 Streamlit 原生的 columns，比例 1:1
col_input, col_preview = st.columns(2, gap="medium")

with col_input:
    st.subheader("⌨️ 编辑区")
    # text_area 设置高度为 600px，足够长
    md_text = st.text_area(
        "输入 Markdown 内容", 
        value=default_text, 
        height=600, 
        label_visibility="collapsed",
        placeholder="在此粘贴您的 Markdown 内容..."
    )

with col_preview:
    st.subheader("👁️ 实时预览")
    
    # 使用 st.container(border=True) 创建一个带边框的容器，替代之前的 CSS hack
    # 这是 Streamlit 新版原生功能，非常稳定
    with st.container(border=True):
        if md_text.strip():
            # 直接使用 Streamlit 内置的 markdown 渲染器
            # 它本身就基于 KaTeX，对 LaTeX 公式支持极好
            st.markdown(md_text, unsafe_allow_html=True)
        else:
            st.info("👈 请在左侧输入内容")

# --- 5. 底部操作栏 ---
st.divider()

# 居中放置下载按钮
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])

with col_btn2:
    # 使用 primary 类型的高亮按钮
    # 逻辑：先点击生成，成功后显示下载按钮
    if st.button("🚀 开始转换并生成 Word 文档", type="primary", use_container_width=True):
        if not md_text.strip():
            st.warning("⚠️ 内容不能为空")
        else:
            with st.spinner("正在调用 Pandoc 引擎进行渲染..."):
                docx_path = convert_to_docx(md_text)
                
            if docx_path and os.path.exists(docx_path):
                # 读取文件二进制数据
                with open(docx_path, "rb") as f:
                    file_data = f.read()
                
                # 显示成功并提供下载
                st.success("✅ 转换成功！点击下方按钮下载。")
                st.download_button(
                    label="⬇️ 下载 Word 文档 (.docx)",
                    data=file_data,
                    file_name="converted_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                # 清理临时文件
                try:
                    os.remove(docx_path)
                except:
                    pass
