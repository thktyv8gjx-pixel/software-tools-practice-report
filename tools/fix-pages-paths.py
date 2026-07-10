from pathlib import Path

PROJECT_PREFIX = "/software-tools-practice-report"

for markdown_file in Path("source").rglob("*.md"):
    # utf-8-sig 同时兼容普通 UTF-8 和带 BOM 的文件
    content = markdown_file.read_text(encoding="utf-8-sig")

    # Markdown 图片路径
    content = content.replace(
        "](/images/",
        f"]({PROJECT_PREFIX}/images/",
    )

    # 兼容正文中可能存在的 HTML img 标签
    content = content.replace(
        'src="/images/',
        f'src="{PROJECT_PREFIX}/images/',
    )

    markdown_file.write_text(content, encoding="utf-8")

print("GitHub Pages image paths updated.")