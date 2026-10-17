"""
PPT 文件批量重命名工具 (v3)

根据 PPT 第一页的标题自动重命名文件。

主要功能:
- 遍历文件夹中的所有 .pptx 文件
- 提取第一张幻灯片的标题或首个文本框内容
- 过滤无效/占位符标题，避免产生垃圾文件名
- 清理并规范化文件名
- 处理文件名冲突
- 生成详细的处理日志
- 双击运行结束后暂停，避免窗口闪退

v3 相对 v2 的改进:
1. 末尾暂停，防止双击运行完成后窗口立即关闭（"闪退"）
2. 过滤空标题、纯符号标题（如 "+"）和未填写的占位符文本
3. 清理文件名首尾句点，避免出现 "Welcome..pptx" 这类结果
4. 标题排序时对形状坐标/尺寸做 None 安全处理

作者: Claude Code
版本: 3.0
"""

import logging
import re
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.exc import PackageNotFoundError
from zipfile import BadZipFile


# ==================== 配置常量 ====================
MAX_FILENAME_LENGTH = 150  # 文件名最大长度（Windows MAX_PATH 为 260）
MIN_TITLE_LENGTH = 2  # 标题最短字符数，低于此长度视为无效
DEFAULT_EXTENSION = ".pptx"
INVALID_CHARS_PATTERN = r'[\\/:*?"<>|]'  # Windows 不允许的文件名字符

# 常见的占位符文本（未填写的标题/文本框的提示文字）
PLACEHOLDER_TEXTS = {
    "click to edit title",
    "click to edit",
    "双击此处添加标题",
    "单击此处添加标题",
    "在此处添加标题",
}


# ==================== 日志配置 ====================
def setup_logging() -> logging.Logger:
    """配置日志系统。"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(message)s',
        handlers=[logging.StreamHandler()]
    )
    return logging.getLogger(__name__)


logger = setup_logging()


# ==================== 核心功能函数 ====================
def clean_filename(text: str) -> str:
    """
    清理字符串，使其成为有效的文件名。

    处理步骤:
    1. 替换 Windows 文件名不允许的字符
    2. 将多个连续的空格/下划线替换为单个下划线
    3. 去除首尾的下划线、空格和句点（Windows 不允许文件名以点/空格结尾）
    4. 限制文件名长度
    """
    if not isinstance(text, str):
        return ""

    # 替换文件名中不允许的字符
    cleaned_text = re.sub(INVALID_CHARS_PATTERN, '_', text)

    # 统一空格和下划线
    cleaned_text = re.sub(r'[\s_]+', '_', cleaned_text)

    # 去除首尾的下划线、空格和句点
    cleaned_text = cleaned_text.strip('_. ')

    # 限制文件名长度
    if len(cleaned_text) > MAX_FILENAME_LENGTH:
        cleaned_text = cleaned_text[:MAX_FILENAME_LENGTH].rstrip('_. ')

    return cleaned_text


def is_valid_title(text: str) -> bool:
    """
    判断文本是否可作为有效的标题。

    排除以下情况:
    - 空文本
    - 过短（少于 MIN_TITLE_LENGTH 个字符）
    - 纯符号文本（不含任何字母、数字或汉字，如 "+"、"---"）
    - 未填写的占位符文本
    """
    if not isinstance(text, str):
        return False

    stripped = text.strip()
    if not stripped:
        return False

    if stripped.lower() in PLACEHOLDER_TEXTS:
        return False

    if len(stripped) < MIN_TITLE_LENGTH:
        return False

    # 必须包含至少一个字母、数字或汉字
    if not re.search(r'[0-9A-Za-z一-鿿]', stripped):
        return False

    return True


def get_unique_filename(filepath: Path, base_name: str, ext: str) -> Path:
    """
    生成唯一的文件名，避免与现有文件冲突。

    如果目标文件名已存在，会自动添加递增的数字后缀。
    """
    counter = 1
    new_filepath = filepath.parent / f"{base_name}{ext}"

    # 如果目标文件已存在且不是原文件本身，添加数字后缀
    while new_filepath.exists() and new_filepath != filepath:
        new_filepath = filepath.parent / f"{base_name}_{counter}{ext}"
        counter += 1

    return new_filepath


def extract_first_slide_title(prs: Presentation) -> Optional[str]:
    """
    从演示文稿中提取第一张幻灯片的标题。

    策略:
    1. 优先使用标题占位符的文本
    2. 如果没有有效标题占位符，查找第一个有效的文本框
    3. 优先选择位置靠上、尺寸较大的文本框（通常是标题）
    """
    if not prs.slides:
        return None

    first_slide = prs.slides[0]

    # 策略1: 尝试获取标题占位符
    if first_slide.shapes.title:
        title_text = first_slide.shapes.title.text.strip()
        if is_valid_title(title_text):
            return title_text

    # 策略2: 查找所有包含有效文本的形状
    text_shapes = []
    for shape in first_slide.shapes:
        if shape.has_text_frame:
            text_content = shape.text_frame.text.strip()
            if is_valid_title(text_content):
                text_shapes.append({
                    'text': text_content,
                    'top': shape.top or 0,
                    'area': (shape.width or 0) * (shape.height or 0)
                })

    if not text_shapes:
        return None

    # 策略3: 按位置（从上到下）和面积（从大到小）排序，选择最可能是标题的文本框
    text_shapes.sort(key=lambda s: (s['top'], -s['area']))
    return text_shapes[0]['text']


def rename_ppts_by_first_slide_title(folder_path: str) -> None:
    """
    遍历指定文件夹中的所有 PPTX 文件，并根据其第一页的标题重命名。
    """
    folder = Path(folder_path)

    logger.info(f"✈️ 开始处理文件夹: {folder}\n")

    if not folder.is_dir():
        logger.error(f"错误: 指定的路径 '{folder}' 不是一个有效的文件夹。")
        return

    # 统计计数器
    stats = {'processed': 0, 'skipped': 0, 'renamed': 0}

    # 遍历所有 .pptx 文件
    for filepath in folder.glob(f"*{DEFAULT_EXTENSION}"):
        # 跳过隐藏文件和系统文件
        if filepath.name.startswith('.'):
            logger.info(f"  ⏭️  跳过隐藏文件: '{filepath.name}'")
            stats['skipped'] += 1
            continue

        logger.info(f"--- 正在处理 '{filepath.name}' ---")

        try:
            # 尝试打开 PPT 文件
            prs = Presentation(str(filepath))

            # 提取第一张幻灯片的标题
            first_slide_title = extract_first_slide_title(prs)

            if not first_slide_title:
                logger.warning(
                    f"  ⚠️  警告: 无法从 '{filepath.name}' 的第一页获取有效的标题或文本内容。"
                )
                stats['skipped'] += 1
                continue

            # 清理标题生成文件名
            cleaned_title = clean_filename(first_slide_title)

            if not cleaned_title:
                logger.warning(
                    f"  ⚠️  警告: 从 '{filepath.name}' 获取的标题 "
                    f"'{first_slide_title}' 清理后为空，跳过。"
                )
                stats['skipped'] += 1
                continue

            # 生成唯一的新文件名
            new_filepath = get_unique_filename(
                filepath,
                cleaned_title,
                filepath.suffix
            )

            # 判断是否需要重命名
            if filepath == new_filepath:
                logger.info(
                    f"  ✅ '{filepath.name}' 的新旧名称相同 "
                    f"('{new_filepath.name}')，无需重命名。"
                )
            else:
                # 执行重命名操作
                filepath.rename(new_filepath)
                logger.info(
                    f"  ✅ 成功重命名 '{filepath.name}' 为 '{new_filepath.name}'"
                )
                stats['renamed'] += 1

            stats['processed'] += 1

        except PackageNotFoundError:
            logger.error(
                f"  ❌ 错误: '{filepath.name}' 不是一个有效的 PowerPoint 文件"
                f"（可能损坏或格式错误）。"
            )
            stats['skipped'] += 1

        except BadZipFile:
            logger.error(
                f"  ❌ 错误: '{filepath.name}' 文件损坏或不是有效的 zip 格式。"
            )
            stats['skipped'] += 1

        except PermissionError:
            logger.error(
                f"  ❌ 错误: 没有权限重命名 '{filepath.name}'，"
                f"文件可能正在被其他程序使用。"
            )
            stats['skipped'] += 1

        except Exception as e:
            logger.error(f"  ❌ 处理 '{filepath.name}' 时发生未知错误: {e}")
            stats['skipped'] += 1

        finally:
            logger.info("-" * (len(filepath.name) + 20) + "\n")

    # 输出统计报告
    logger.info("✨ 所有 PPT 文档处理完毕！")
    logger.info(
        f"统计：已处理 {stats['processed']} 个文件，"
        f"成功重命名 {stats['renamed']} 个，跳过 {stats['skipped']} 个。"
    )


# ==================== 主程序入口 ====================
if __name__ == "__main__":
    # 自动获取脚本所在的当前文件夹的路径
    script_dir = Path(__file__).parent.absolute()

    # 运行重命名程序
    rename_ppts_by_first_slide_title(str(script_dir))

    # 双击运行时暂停，避免窗口立即关闭（命令行/管道环境会自动跳过）
    try:
        input("\n按回车键退出...")
    except (EOFError, KeyboardInterrupt):
        pass
