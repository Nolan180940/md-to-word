"""
Markdown → Word 桌面版 (customtkinter 现代化 GUI)
原生 Windows 窗口应用，粘贴 Markdown → 一键生成精美 DOCX
"""

import sys
import os
import re
import tempfile
import threading
import shutil
import tkinter.filedialog
import tkinter.messagebox

import pypandoc
import customtkinter as ctk

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ============================================================
# 捆绑 Pandoc 路径处理
# ============================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

pandoc_dir = os.path.join(BASE_DIR, 'pandoc')
os.environ['PATH'] = pandoc_dir + os.pathsep + os.environ.get('PATH', '')
os.environ['PYPANDOC_PANDOC'] = os.path.join(pandoc_dir, 'pandoc.exe')

# ============================================================
# 核心功能
# ============================================================

def smart_fix_markdown(text):
    log = []
    fixed_text = text if text is not None else ""

    if '\u200b' in fixed_text:
        fixed_text = fixed_text.replace('\u200b', '')
        log.append("🧹 移除了隐形字符")

    if '\\[' in fixed_text or '\\]' in fixed_text:
        fixed_text = fixed_text.replace('\\[', '$$').replace('\\]', '$$')
        log.append("📐 LaTeX 块级公式 \\[...\\] → $$...$$")

    if '\\(' in fixed_text or '\\)' in fixed_text:
        fixed_text = fixed_text.replace('\\(', '$').replace('\\)', '$')
        log.append("📐 LaTeX 行内公式 \\(...\\) → $...$")

    pattern_space_math = r'(?<!\$)\$[ \t]+(.*?)[ \t]+\$(?!\$)'
    if re.search(pattern_space_math, fixed_text):
        new_text, count = re.subn(pattern_space_math, r'$\1$', fixed_text)
        if count > 0:
            fixed_text = new_text
            log.append(f"🔧 移除 {count} 处行内公式多余空格")

    if '<sup>' in fixed_text:
        new_text, count = re.subn(r'<sup>(.*?)</sup>', r'^\1^', fixed_text)
        if count > 0:
            fixed_text = new_text
            log.append(f"⬆️ 将 {count} 处 HTML 上标转为 Markdown")

    code_fence_count = len(re.findall(r'^```', fixed_text, re.MULTILINE))
    if code_fence_count % 2 != 0:
        fixed_text += "\n```"
        log.append("🧱 自动闭合未结束的代码块")

    math_block_count = fixed_text.count('$$')
    if math_block_count % 2 != 0:
        fixed_text += "\n$$"
        log.append("🧮 自动闭合未结束的公式块")

    fixed_text = re.sub(r'([^\n])\n```', r'\1\n\n```', fixed_text)
    fixed_text = re.sub(r'```\n([^\n])', r'```\n\n\1', fixed_text)

    lines = fixed_text.splitlines()
    if len(lines) == 0:
        return fixed_text, log

    out_lines = []
    i = 0
    made_change = False
    while i < len(lines):
        line = lines[i]
        stripped = line.lstrip()
        if stripped.startswith('>'):
            if out_lines and out_lines[-1].strip() != '':
                out_lines.append('')
                made_change = True
            while i < len(lines) and lines[i].lstrip().startswith('>'):
                out_lines.append(lines[i])
                i += 1
            if i < len(lines) and lines[i].strip() != '':
                out_lines.append('')
                made_change = True
            continue
        else:
            out_lines.append(line)
            i += 1

    ends_with_newline = fixed_text.endswith('\n')
    new_fixed = '\n'.join(out_lines)
    if ends_with_newline:
        new_fixed = new_fixed + '\n'
    if made_change:
        fixed_text = new_fixed
        log.append("🧩 blockquote 段落前后强制空行")

    return fixed_text, log


def apply_word_styles(docx_path):
    if not HAS_DOCX:
        return
    doc = Document(docx_path)
    styles = doc.styles
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
    except Exception:
        pass
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
    except Exception:
        pass
    doc.save(docx_path)


def convert_to_docx(md_content):
    output_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            output_path = tmp_file.name
        pypandoc.convert_text(
            md_content, 'docx',
            format='markdown-yaml_metadata_block+tex_math_dollars',
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
            except Exception:
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


# ============================================================
DEFAULT_TEXT = r'''# 深度学习中的概率分布

这是一个包含 "空格公式" 的测试。

## 1. 坏掉的公式 (Spaces)

大模型经常输出这种带空格的行内公式： $E = mc^2$ ，或者 $ x_0 = 0 $。
本工具会自动将其修复为：$E=mc^2$ 和 $x_0=0$。

## 2. 块级公式 (LaTeX 风格)

\[
\mathcal{L}(\theta) = -\frac{1}{N} \sum_{i=1}^N \left[ y_i \log(\hat{y}_i) + (1-y_i) \log(1-\hat{y}_i) \right]
\]

## 3. 代码块测试

```python
def fix_spaces(text):
    return text.strip()
```
'''


# ============================================================
# 现代化 GUI 应用 (customtkinter)
# ============================================================

class ModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Markdown → Word 转换器")
        self.geometry("1000x720")
        self.minsize(700, 500)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - 1000) // 2
        y = (sh - 720) // 2
        self.geometry(f"1000x720+{x}+{y}")

        self._temp_docx = None
        self._is_working = False
        self._build_ui()
        self._bind_keys()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 顶部标题栏
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 0))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="📝 Markdown → Word 转换器",
                     font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text="粘贴 Markdown 内容，一键生成精美排版的 Word 文档",
                     font=ctk.CTkFont(size=12), text_color="gray").grid(
                         row=1, column=0, sticky="w", pady=(2, 0))

        self._theme_btn = ctk.CTkButton(
            header, text="🌙", width=40, height=32,
            font=ctk.CTkFont(size=16),
            fg_color="transparent", hover_color=("gray80", "gray30"),
            command=self._toggle_theme)
        self._theme_btn.grid(row=0, column=1, rowspan=2, sticky="e", padx=(8, 0))
        self._update_theme_btn()

        # 工具栏
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=28, pady=(16, 6))
        toolbar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(toolbar, text="✏️  Markdown 编辑区",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w")

        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.grid(row=0, column=1, sticky="e")
        ctk.CTkButton(btn_frame, text="📂 打开文件", width=100, height=32,
                       font=ctk.CTkFont(size=12),
                       fg_color="transparent", border_width=1,
                       command=self._load_file).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="🔄 重置示例", width=100, height=32,
                       font=ctk.CTkFont(size=12),
                       fg_color="transparent", border_width=1,
                       command=self._reset_default).pack(side="left")

        # 文本编辑区
        self.text_area = ctk.CTkTextbox(
            self, font=ctk.CTkFont(family="Consolas", size=13),
            wrap="word", border_width=1, corner_radius=10,
            activate_scrollbars=True)
        self.text_area.grid(row=2, column=0, sticky="nsew", padx=28)
        self.text_area.insert("1.0", DEFAULT_TEXT)

        # 修复日志区（初始隐藏）
        self.log_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.log_header = ctk.CTkLabel(self.log_frame, text="",
                                        font=ctk.CTkFont(size=12, weight="bold"),
                                        text_color=("#2d6a4f", "#52b788"))
        self.log_text = ctk.CTkTextbox(
            self.log_frame, height=70,
            font=ctk.CTkFont(size=11), wrap="word",
            border_width=0, corner_radius=8,
            fg_color=("#f0fdf4", "#0a3622"),
            text_color=("#166534", "#d8f3dc"),
            activate_scrollbars=True)
        self.log_text.configure(state="disabled")

        # 底部操作栏
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", padx=28, pady=(8, 20))
        footer.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(
            footer, text="", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"))
        self.status_label.grid(row=0, column=0, sticky="w", padx=(4, 0))

        self.convert_btn = ctk.CTkButton(
            footer, text="🚀 生成并保存 Word 文档",
            height=44, font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=10, command=self._convert_and_save)
        self.convert_btn.grid(row=0, column=1, sticky="e")

    def _bind_keys(self):
        self.bind('<Control-Return>', lambda e: self._convert_and_save())
        self.bind('<Control-o>', lambda e: self._load_file())

    def _toggle_theme(self):
        cur = ctk.get_appearance_mode()
        ctk.set_appearance_mode("Dark" if cur == "Light" else "Light")
        self._update_theme_btn()

    def _update_theme_btn(self):
        self._theme_btn.configure(text="☀️" if ctk.get_appearance_mode() == "Dark" else "🌙")

    def _reset_default(self):
        self.text_area.delete("1.0", "end")
        self.text_area.insert("1.0", DEFAULT_TEXT)
        self._clear_status()

    def _load_file(self):
        fp = tkinter.filedialog.askopenfilename(
            title="打开 Markdown 文件",
            filetypes=[("Markdown 文件", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")])
        if fp:
            try:
                with open(fp, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", content)
                self._clear_status()
                self._set_status(f"✅ 已加载：{os.path.basename(fp)}")
            except Exception as e:
                tkinter.messagebox.showerror("读取失败", f"无法读取文件：\n{e}")

    def _set_status(self, msg, is_error=False):
        self.status_label.configure(
            text=msg,
            text_color=("#e03131", "#ff6b6b") if is_error else ("gray50", "gray60"))

    def _clear_status(self):
        self.status_label.configure(text="")
        self._hide_log()

    def _show_log(self, logs):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        for item in logs:
            self.log_text.insert("end", f"  {item}\n")
        self.log_text.configure(state="disabled")
        self.log_header.configure(text=f"🤖 自动执行了 {len(logs)} 项智能修复")
        self.log_header.grid(row=0, column=0, sticky="w", padx=4, pady=(8, 4))
        self.log_text.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
        self.log_frame.grid(row=3, column=0, sticky="ew", padx=28, pady=(8, 0))

    def _hide_log(self):
        self.log_frame.grid_forget()

    def _convert_and_save(self):
        if self._is_working:
            return
        md_text = self.text_area.get("1.0", "end-1c").strip()
        if not md_text:
            tkinter.messagebox.showwarning("内容为空", "请先输入或粘贴 Markdown 内容。")
            return

        self._is_working = True
        self.convert_btn.configure(text="⏳ 正在转换...", state="disabled",
                                    fg_color=("gray60", "gray50"))
        self._set_status("正在处理中...")
        self.update_idletasks()

        def do_work():
            fixed_text, logs = smart_fix_markdown(md_text)
            filename = generate_smart_filename(fixed_text)
            docx_path, error = convert_to_docx(fixed_text)
            self.after(0, lambda: self._on_convert_done(docx_path, error, filename, logs))

        threading.Thread(target=do_work, daemon=True).start()

    def _on_convert_done(self, docx_path, error, filename, logs):
        self._is_working = False
        self.convert_btn.configure(text="🚀 生成并保存 Word 文档", state="normal",
                                    fg_color=("#3b82f6", "#1d4ed8"))

        if error:
            self._set_status(f"❌ 转换失败：{error}", is_error=True)
            tkinter.messagebox.showerror("转换失败", error)
            return

        if logs:
            self._show_log(logs)

        save_path = tkinter.filedialog.asksaveasfilename(
            title="保存 Word 文档", defaultextension=".docx",
            initialfile=filename,
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])

        if save_path:
            try:
                shutil.copy2(docx_path, save_path)
                self._set_status(f"✅ 已保存：{os.path.basename(save_path)}")
            except Exception as e:
                self._set_status(f"❌ 保存失败：{e}", is_error=True)
                tkinter.messagebox.showerror("保存失败", str(e))
        else:
            self._set_status("已取消保存")

        if docx_path and os.path.exists(docx_path):
            try:
                os.remove(docx_path)
            except Exception:
                pass

    def destroy(self):
        if self._temp_docx and os.path.exists(self._temp_docx):
            try:
                os.remove(self._temp_docx)
            except Exception:
                pass
        super().destroy()


def main():
    app = ModernApp()
    app.mainloop()


if __name__ == '__main__':
    main()
