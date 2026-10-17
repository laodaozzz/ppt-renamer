# ppt-renamer

[English](#english) | [中文](#中文)

---

## English

A lightweight Python tool that batch-renames PowerPoint (`.pptx`) files using the title of their first slide.

### Features

- Scans a folder for all `.pptx` files
- Extracts the first slide's title (from the title placeholder, or falls back to the topmost / largest text box)
- Filters out empty, placeholder, or invalid titles (e.g. `+`)
- Cleans filenames: removes Windows-invalid characters, collapses whitespace/underscores, trims trailing dots and underscores
- Handles filename collisions by appending a numeric suffix
- Prints detailed per-file logs and a summary
- Pauses before exit so the console window stays open when double-clicked

### Installation

Requires Python 3 and `python-pptx`:

```bash
pip install python-pptx
```

### Usage

Place `PPTRenamer3.py` in the folder containing your `.pptx` files and run it — either double-click it, or:

```bash
python PPTRenamer3.py
```

The script processes all `.pptx` files in its own folder and renames each one to its first-slide title.

### Example

| Before | After |
|--------|-------|
| `lecture-slide-2024.pptx` | `Enhanced_Recovery_After_Surgery.pptx` |
| `presentation (1).pptx` | `Welcome.pptx` |

---

## 中文

一个轻量级 Python 工具，根据 PPT 第一页的标题批量重命名 `.pptx` 文件。

### 功能特点

- 扫描文件夹中所有 `.pptx` 文件
- 提取第一张幻灯片的标题（优先标题占位符，否则取最靠上 / 最大的文本框）
- 过滤空标题、占位符文本和无效标题（如 `+`）
- 清理文件名：移除 Windows 非法字符、合并空格 / 下划线、去除首尾句点和下划线
- 文件名冲突时自动添加数字后缀
- 输出详细的逐文件日志和统计报告
- 运行结束自动暂停，双击运行时窗口不会闪退

### 安装

需要 Python 3 和 `python-pptx`：

```bash
pip install python-pptx
```

### 使用

将 `PPTRenamer3.py` 放到包含 `.pptx` 文件的文件夹中，双击运行，或在命令行执行：

```bash
python PPTRenamer3.py
```

脚本会处理其所在文件夹中的所有 `.pptx` 文件，并按第一页标题重命名。

### 示例

| 重命名前 | 重命名后 |
|---------|---------|
| `lecture-slide-2024.pptx` | `Enhanced_Recovery_After_Surgery.pptx` |
| `presentation (1).pptx` | `Welcome.pptx` |

---

## License

MIT
