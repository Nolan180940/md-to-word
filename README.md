# 🛠️ Markdown to Word Pro

一个智能的 Markdown 转 Word 转换器，专为处理 LLM 生成的内容而设计。支持 LaTeX 公式渲染、自动语法修复和精美的 Word 样式输出。

![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?logo=streamlit&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-Open%20Source-green.svg)

## ✨ 特性亮点

- **📝 Markdown → Word 转换**：高质量转换为可编辑的 `.docx` 格式
- **🧮 LaTeX 公式支持**：完美渲染行内公式 `$...$` 和块级公式 `$$...$$`
- **🤖 LLM 内容优化**：自动修复 AI 生成的不标准 Markdown 语法
- **🎨 精美样式输出**：代码块带阴影边框，引用块带左侧缩进线
- **📄 智能文件名**：根据文档标题自动生成合适的文件名
- **🆓 免费开源**：无需登录，无限使用，完全开放源码

## 🚀 快速开始

### 在线使用

访问部署实例即可直接使用（如有部署链接请在此添加）。

### 本地运行

#### 系统要求

- Python 3.8+
- Pandoc（系统级安装）

#### 安装步骤

1. **克隆仓库**
   ```bash
   git clone <repository-url>
   cd md-to-word
   ```

2. **安装 Pandoc**

   - **Ubuntu/Debian**:
     ```bash
     sudo apt-get install pandoc
     ```
   
   - **macOS**:
     ```bash
     brew install pandoc
     ```
   
   - **Windows**:
     从 [Pandoc 官网](https://pandoc.org/installing.html) 下载安装

3. **安装 Python 依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **启动应用**
   ```bash
   streamlit run app.py
   ```

5. **访问应用**

   浏览器打开 `http://localhost:8501`

## 📋 依赖说明

| 文件 | 用途 |
|------|------|
| `requirements.txt` | Python 包依赖 (streamlit, pypandoc, python-docx) |
| `packages.txt` | 系统级依赖 (pandoc) |

## 🔧 核心功能详解

### 1. 智能 Markdown 修复

自动检测并修复以下常见问题：

| 问题类型 | 修复内容 |
|----------|----------|
| 隐形字符 | 移除零宽空格 (`\u200b`) |
| LaTeX 语法 | `\[..., \]` → `$$...$$`，`\(...\)` → `$...$` |
| 公式空格 | `$ x^2 $` → `$x^2$` |
| HTML 标签 | `<sup>...</sup>` → `^...^` |
| 未闭合代码块 | 自动添加缺失的 ```` ``` ```` |
| 未闭合公式 | 自动添加缺失的 `$$` |
| Blockquote 格式 | 确保引用块前后有空行 |

### 2. Word 样式增强

转换后的文档包含专业样式：

- **代码块**：Consolas 字体 + 灰色背景 + 四边边框
- **引用块**：左侧缩进 + 灰色竖线 + 正体显示
- **数学公式**：保持 LaTeX 渲染效果

### 3. 实时预览

左侧编辑，右侧实时预览修复后的 Markdown 效果，所见即所得。

## 💡 使用示例

在输入框中粘贴任意 Markdown 内容，例如：

```markdown
# 深度学习中的概率分布

## 行内公式

大模型经常输出这种带空格的行内公式：$E = mc^2$，本工具会自动修复为：$E=mc^2$

## 块级公式

\[
\mathcal{L}(\theta) = -\frac{1}{N} \sum_{i=1}^N y_i \log(\hat{y}_i)
\]

## 代码示例

```python
def hello():
    print("Hello, World!")
```


点击「生成定制化 Word 文档」按钮，即可获得格式精美的 Word 文件。

## 🏗️ 项目结构

```
.
├── app.py              # 主应用程序（Streamlit）
├── requirements.txt    # Python 依赖
├── packages.txt        # 系统依赖
└── README.md           # 项目文档
```

## 🛠️ 技术栈

- **前端框架**: [Streamlit](https://streamlit.io/)
- **转换引擎**: [Pandoc](https://pandoc.org/) + [pypandoc](https://github.com/bebraw/pypandoc)
- **文档处理**: [python-docx](https://python-docx.readthedocs.io/)

## 📝 License

本项目为开源软件，自由使用、修改和分发。

## 🙏 致谢

感谢以下开源项目：
- Streamlit 团队提供的优秀框架
- Pandoc 文档转换工具
- python-docx 库的样式处理能力

---

**享受无缝的 Markdown 到 Word 转换体验！** 🎉
