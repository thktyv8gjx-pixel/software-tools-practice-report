import html
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DOCX_DIR = ROOT / "input-docx"
OUT_DIR = ROOT / "source" / "report-demo"
IMAGE_DIR = OUT_DIR / "assets" / "images"
DOCX = DOCX_DIR / "软件开发工具实践实验报告_任垚橙_25051414_最终版.docx"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


SELECTED_IMAGES = [
    ("image1.png", "architecture-toolchain.png", "双通道工具链总览", "1.2 目标场景", "项目与代码库配置"),
    ("image2.png", "network-topology.png", "NAT 与 Host-Only 网络拓扑", "2.3 网络拓扑与地址规划", "Linux 环境配置"),
    ("image3.png", "tool-version-check.png", "Node.js、npm、Git、VS Code 与 Hexo 版本验证", "3.1 实验一", "Web 运行环境"),
    ("image4.png", "vscode-hexo-project.png", "VS Code 中的 Hexo 项目结构", "3.1 实验一", "项目与代码库配置"),
    ("image10.png", "markdown-preview.png", "Markdown 源码与实时预览对照检查", "3.2 实验二", "项目与代码库配置"),
    ("image11.png", "git-status-config.png", "Git 用户信息配置与项目状态", "3.2 实验二", "项目与代码库配置"),
    ("image12.png", "git-commits.png", "Git 提交记录", "3.2 实验二", "项目与代码库配置"),
    ("image14.png", "github-repository.png", "GitHub 远程仓库文件", "3.2 实验二", "项目与代码库配置"),
    ("image16.png", "virtualbox-hardware.png", "Ubuntu Server 虚拟机硬件配置", "3.3 实验三", "Linux 环境配置"),
    ("image17.png", "virtualbox-network.png", "VirtualBox NAT 与 Host-Only 双网卡", "3.3 实验三", "Linux 环境配置"),
    ("image21.png", "virtualbox-nodes.png", "VirtualBox 完整克隆三个 Ubuntu 节点", "3.3 实验三", "Linux 环境配置"),
    ("image22.png", "node1-ssh-ip.png", "node1 主机名、静态 IP 与 SSH 服务验证", "3.3 实验三", "Linux 环境配置"),
    ("image25.png", "nodes-connectivity.png", "node1 与 node2、node3 网络互联验证", "3.3 实验三", "Linux 环境配置"),
    ("image34.png", "bashrc-env-alias.png", "~/.bashrc 环境变量与命令别名", "3.6 实验六", "Linux 环境配置"),
    ("image37.png", "env-persistence.png", "重新登录后环境变量与自定义命令持久有效", "3.6 实验六", "Linux 环境配置"),
    ("image38.png", "scp-upload.png", "Windows 通过 SCP 上传测试文件", "3.7 实验七", "远程管理配置"),
    ("image41.png", "ssh-config-login.png", "SSH 客户端 Config 快捷免密登录", "3.7 实验七", "远程管理配置"),
    ("image42.png", "ubuntu-bare-git.png", "Ubuntu 裸 Git 仓库推送验证", "3.7 实验七", "远程管理配置"),
    ("image46.png", "rsync-delete-sync.png", "rsync 增量同步与 --delete 删除验证", "3.7 实验七", "远程管理配置"),
    ("image47.png", "nginx-running.png", "Nginx 安装、开机自启与运行状态", "3.8 实验八", "Web 站点部署"),
    ("image48.png", "nginx-config-test.png", "Nginx 站点配置语法检查成功", "3.8 实验八", "Web 站点部署"),
    ("image50.png", "private-site-browser.png", "通过虚拟机 IP 访问实验报告站点", "3.8 实验八", "Web 站点部署"),
    ("image52.png", "deploy-script-success.png", "自动部署完成并推送两个远程仓库", "3.8 实验八", "Web 站点部署"),
    ("image53.png", "github-pages-settings.png", "GitHub Pages 发布来源设置", "3.8 实验八", "Web 站点部署"),
    ("image54.png", "github-actions-success.png", "GitHub Actions 部署成功", "3.8 实验八", "Web 站点部署"),
    ("image55.png", "github-pages-final.png", "GitHub Pages 公网完整样式访问成功", "3.8 实验八", "Web 站点部署"),
    ("image56.jpeg", "node-command-error.jpeg", "Node.js 命令无法识别", "6.1 典型问题", "Web 运行环境"),
    ("image57.jpeg", "powershell-npm-policy.jpeg", "PowerShell 阻止 npm.ps1", "6.2 典型问题", "Web 运行环境"),
    ("image58.jpeg", "localhost-refused.jpeg", "localhost 页面拒绝连接", "6.3 典型问题", "Web 运行环境"),
    ("image59.jpeg", "virtualbox-boot-medium.jpeg", "虚拟机无法读取启动介质", "6.4 典型问题", "Linux 环境配置"),
    ("image60.jpeg", "grub-install-error.jpeg", "Ubuntu 安装在 GRUB 阶段失败", "6.5 典型问题", "Linux 环境配置"),
    ("image61.jpeg", "netplan-yaml-error.jpeg", "Netplan YAML 解析失败", "6.6 典型问题", "Linux 环境配置"),
    ("image62.jpeg", "ssh-host-alias-error.jpeg", "SSH 快捷主机名无法解析", "6.7 典型问题", "远程管理配置"),
    ("image63.jpeg", "hexo-powershell-script-error.jpeg", "Hexo 误加载 PowerShell 部署脚本", "6.8 典型问题", "Web 站点部署"),
    ("image64.png", "disk-full-error.png", "VirtualBox 克隆时磁盘空间不足", "6.9 典型问题", "Linux 环境配置"),
    ("image65.png", "clone-ssh-timeout.png", "克隆节点启动卡住与 SSH 超时", "6.10 典型问题", "远程管理配置"),
    ("image66.png", "github-pages-css-404.png", "GitHub Pages 主题 CSS 返回 404", "6.11 典型问题", "Web 站点部署"),
    ("image67.png", "ci-container-upgrade.png", "由虚拟机手工部署升级为容器化持续部署", "8.1 技术展望", "Web 站点部署"),
]


def qn(prefix, tag):
    return f"{{{NS[prefix]}}}{tag}"


def text_of(element):
    parts = []
    for node in element.iter():
        if node.tag == qn("w", "t") and node.text:
            parts.append(node.text)
        elif node.tag == qn("w", "tab"):
            parts.append("\t")
        elif node.tag == qn("w", "br"):
            parts.append("\n")
    return "".join(parts).strip()


def read_docx_metadata():
    if not DOCX.exists():
        raise FileNotFoundError(DOCX)

    with zipfile.ZipFile(DOCX) as archive:
        media_files = [name for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")]
        styles_root = ET.fromstring(archive.read("word/styles.xml"))
        style_map = {}
        for style in styles_root.findall("w:style", NS):
            style_id = style.get(qn("w", "styleId"))
            name_el = style.find("w:name", NS)
            style_map[style_id] = name_el.get(qn("w", "val")) if name_el is not None else style_id

        rel_root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        rels = {
            rel.get("Id"): rel.get("Target")
            for rel in rel_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        }

        root = ET.fromstring(archive.read("word/document.xml"))
        body = root.find("w:body", NS)
        paragraphs = []
        headings = []
        top_headings = []
        captions = []
        image_refs = []
        tables = []

        block_index = 0
        for child in list(body):
            if child.tag == qn("w", "p"):
                block_index += 1
                text = text_of(child)
                style_name = ""
                style_id = None
                ppr = child.find("w:pPr", NS)
                if ppr is not None:
                    pstyle = ppr.find("w:pStyle", NS)
                    if pstyle is not None:
                        style_id = pstyle.get(qn("w", "val"))
                        style_name = style_map.get(style_id, style_id)

                if text:
                    paragraphs.append(text)
                    is_heading = style_id and (
                        style_id.lower().startswith("heading")
                        or "heading" in style_name.lower()
                        or "标题" in style_name
                        or re.match(r"^\d+(?:\.\d+)*\s+", text)
                    )
                    if is_heading:
                        headings.append(text)
                        if style_name == "heading 1":
                            top_headings.append(text)
                    if text.startswith("图") or text.startswith("Figure"):
                        captions.append(text)

                for blip in child.findall(".//a:blip", NS):
                    rel_id = blip.get(qn("r", "embed")) or blip.get(qn("r", "link"))
                    if rel_id:
                        image_refs.append(rels.get(rel_id, rel_id))

            elif child.tag == qn("w", "tbl"):
                rows = []
                for tr in child.findall("w:tr", NS):
                    row = [text_of(tc) for tc in tr.findall("w:tc", NS)]
                    rows.append(row)
                tables.append(rows)

        caption_by_media = map_captions(archive, rels, style_map)

    return {
        "docx": str(DOCX),
        "paragraphs": len(paragraphs),
        "headings": len(headings),
        "top_headings": len(top_headings),
        "tables": len(tables),
        "media": len(media_files),
        "image_refs": len(image_refs),
        "captions": len(captions),
        "caption_by_media": caption_by_media,
    }


def map_captions(archive, rels, style_map):
    root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    sequence = []
    current_heading = ""

    def style_for(paragraph):
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            return "", None
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is None:
            return "", None
        style_id = pstyle.get(qn("w", "val"))
        return style_map.get(style_id, style_id), style_id

    def add_paragraph(paragraph):
        nonlocal current_heading
        text = text_of(paragraph)
        style_name, style_id = style_for(paragraph)
        if text:
            is_heading = style_id and ("heading" in style_name.lower() or style_id.lower().startswith("heading"))
            if is_heading:
                current_heading = text
            sequence.append({"kind": "text", "text": text, "heading": current_heading})
        for blip in paragraph.findall(".//a:blip", NS):
            rel_id = blip.get(qn("r", "embed")) or blip.get(qn("r", "link"))
            target = rels.get(rel_id)
            if target:
                sequence.append({"kind": "image", "target": target, "heading": current_heading})

    for child in list(body):
        if child.tag == qn("w", "p"):
            add_paragraph(child)
        elif child.tag == qn("w", "tbl"):
            for tc in child.findall(".//w:tc", NS):
                for paragraph in tc.findall("w:p", NS):
                    add_paragraph(paragraph)

    caption_by_media = {}
    for index, item in enumerate(sequence):
        if item["kind"] != "image":
            continue
        caption = ""
        for next_item in sequence[index + 1 : index + 8]:
            if next_item["kind"] == "text":
                if not caption:
                    caption = next_item["text"]
                if next_item["text"].startswith("图"):
                    caption = next_item["text"]
                    break
        caption_by_media[Path(item["target"]).name] = {
            "caption": caption,
            "heading": item.get("heading", ""),
        }
    return caption_by_media


def copy_selected_images():
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(DOCX) as archive:
        media_names = set(archive.namelist())
        for source_name, output_name, *_ in SELECTED_IMAGES:
            zipped = f"word/media/{source_name}"
            if zipped not in media_names:
                raise FileNotFoundError(f"missing media file in DOCX: {zipped}")
            target = IMAGE_DIR / output_name
            with archive.open(zipped) as src, target.open("wb") as dst:
                shutil.copyfileobj(src, dst)


def e(value):
    return html.escape(str(value), quote=True)


def code_block(title, language, code):
    return f"""
            <article class="code-card reveal">
              <div class="code-card__header">
                <span>{e(title)}</span>
                <button class="copy-btn" type="button" data-copy aria-label="复制{e(title)}代码">复制</button>
              </div>
              <pre class="code-block" data-language="{e(language)}"><code>{e(code.strip())}</code></pre>
            </article>"""


def figure_card(image_name, title, caption, experiment, assessment, index):
    alt = f"{title}，{caption}"
    return f"""
            <figure class="gallery-card reveal" data-experiment="{e(experiment)}" data-assessment="{e(assessment)}">
              <button class="gallery-card__media lightbox-trigger" type="button" data-index="{index}" aria-label="放大查看：{e(title)}">
                <img src="./assets/images/{e(image_name)}" alt="{e(alt)}" loading="lazy">
              </button>
              <figcaption>
                <strong>{e(title)}</strong>
                <span>{e(caption)}</span>
                <small>{e(experiment)} · {e(assessment)}</small>
              </figcaption>
            </figure>"""


def build_html(stats):
    caption_by_media = stats["caption_by_media"]
    selected_figures = []
    for source_name, output_name, title, experiment, assessment in SELECTED_IMAGES:
        mapped = caption_by_media.get(source_name, {})
        selected_figures.append({
            "source": source_name,
            "output": output_name,
            "title": title,
            "caption": mapped.get("caption") or title,
            "experiment": experiment,
            "assessment": assessment,
        })

    gallery_html = "\n".join(
        figure_card(fig["output"], fig["title"], fig["caption"], fig["experiment"], fig["assessment"], idx)
        for idx, fig in enumerate(selected_figures)
    )

    command_blocks = [
        ("工具链版本验证", "PowerShell", """node -v
npm -v
git --version
code --version
npm config set registry https://registry.npmmirror.com
npm config get registry
npm install -g hexo-cli
hexo -v"""),
        ("Hexo 项目初始化与本地预览", "PowerShell", """mkdir SoftwareToolsPractice
cd SoftwareToolsPractice
hexo init report-site
cd report-site
npm install
hexo server
hexo new "软件开发工具实践实验报告" """),
        ("Git 本地仓库与 GitHub 远程", "Git", """git config --global user.name "任垚橙"
git config --global user.email "个人邮箱"
git init
git add .
git commit -m "项目初始化并完成Web开发环境配置"
git branch -M main
git remote add origin https://github.com/thktyv8gjx-pixel/software-tools-practice-report.git
git push -u origin main"""),
        ("Netplan Host-Only 静态地址", "YAML", """network:
  version: 2
  ethernets:
    enp0s3:
      dhcp4: true
    enp0s8:
      dhcp4: false
      addresses:
        - 192.168.56.101/24"""),
        ("克隆节点独立身份配置", "Bash", """sudo hostnamectl set-hostname node1
# node2、node3分别设置对应主机名
sudo rm -f /etc/machine-id
sudo systemd-machine-id-setup
sudo rm -f /etc/ssh/ssh_host_*
sudo ssh-keygen -A
sudo systemctl enable --now ssh"""),
        ("共享目录账户与权限", "Bash", """sudo groupadd webgroup
sudo usermod -aG webgroup renyaocheng
sudo usermod -aG webgroup webadmin
sudo mkdir -p /srv/software-tools-site
sudo chown root:webgroup /srv/software-tools-site
sudo chmod 2770 /srv/software-tools-site
ls -ld /srv/software-tools-site"""),
        ("用户级环境变量与别名", "Bash", """export PROJECT_HOME="$HOME/software-tools-lab"
export SITE_HOME="/srv/software-tools-site"
export PATH="$HOME/bin:$PATH"
alias cproject='cd "$PROJECT_HOME"'
alias csite='cd "$SITE_HOME"'
alias ll='ls -alF'
alias labip='ip -br addr'"""),
        ("SSH 客户端快捷主机", "SSH Config", """Host ubuntu-lab
    HostName 192.168.56.101
    User renyaocheng
    Port 22
    IdentityFile ~/.ssh/id_ed25519
    ServerAliveInterval 60"""),
        ("SCP 与 rsync 远程传输", "Bash", """scp ./transfer/windows-upload.txt renyaocheng@192.168.56.101:/home/renyaocheng/
scp test.txt renyaocheng@node2:/home/renyaocheng/
scp renyaocheng@node2:/home/renyaocheng/test.txt ./download-test.txt
rsync -av --delete ./sync-source/ renyaocheng@node2:/home/renyaocheng/sync-target/"""),
        ("Nginx 静态站点配置", "Nginx", """server {
    listen 80;
    listen [::]:80;
    server_name _;
    root /srv/software-tools-site;
    index index.html;
    location / {
        try_files $uri $uri/ =404;
    }
}"""),
        ("PowerShell 自动部署流程", "PowerShell", """$ErrorActionPreference = "Stop"
hexo clean
hexo generate
ssh ubuntu-lab "rm -rf ~/public"
scp -r ./public ubuntu-lab:/home/renyaocheng/
ssh ubuntu-lab "rsync -av --delete ~/public/ /srv/software-tools-site/ && curl -I http://localhost"
Write-Host "Deployment completed: http://192.168.56.101" """),
        ("GitHub Actions 发布摘要", "YAML", """on:
  push:
    branches: [main]
permissions:
  contents: write
# npm ci -> hexo generate -> 验证public -> 发布gh-pages"""),
    ]
    commands_html = "\n".join(code_block(*block) for block in command_blocks)

    return f"""---
layout: false
---
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>软件开发工具实践实验报告成果演示</title>
  <meta name="description" content="任垚橙软件开发工具实践项目成果可视化演示网页">
  <link rel="stylesheet" href="./style.css">
</head>
<body>
  <a class="skip-link" href="#main">跳转到主要内容</a>
  <header class="site-header" id="top">
    <nav class="nav" aria-label="主导航">
      <a class="nav__brand" href="#home" aria-label="返回首页">
        <span class="brand-mark" aria-hidden="true"></span>
        <span>软件开发工具实践</span>
      </a>
      <button class="nav__toggle" type="button" aria-label="展开导航菜单" aria-expanded="false" aria-controls="nav-menu">
        <span></span><span></span><span></span>
      </button>
      <ul class="nav__menu" id="nav-menu">
        <li><a href="#home">首页</a></li>
        <li><a href="#overview">项目概览</a></li>
        <li><a href="#architecture">技术架构</a></li>
        <li><a href="#assessment">考核内容</a></li>
        <li><a href="#process">实验过程</a></li>
        <li><a href="#nodes">多节点环境</a></li>
        <li><a href="#issues">问题解决</a></li>
        <li><a href="#outcomes">成果展示</a></li>
        <li><a href="#summary">总结</a></li>
      </ul>
    </nav>
  </header>

  <main id="main">
    <section class="hero section" id="home" aria-labelledby="hero-title">
      <div class="hero__grid" aria-hidden="true"></div>
      <div class="hero__content reveal">
        <p class="eyebrow">杭州电子科技大学 · 计算机学院</p>
        <h1 id="hero-title">软件开发工具实践实验报告</h1>
        <p class="hero__subtitle">基于 Hexo、Git、Ubuntu Server 与 Nginx 的静态站点开发和部署</p>
        <dl class="hero-meta" aria-label="报告封面信息">
          <div><dt>学校</dt><dd>杭州电子科技大学</dd></div>
          <div><dt>学院</dt><dd>计算机学院</dd></div>
          <div><dt>专业班级</dt><dd>软件工程3班（25052713）</dd></div>
          <div><dt>姓名</dt><dd>任垚橙</dd></div>
          <div><dt>学号</dt><dd>25051414</dd></div>
          <div><dt>完成日期</dt><dd>2026年7月10日</dd></div>
        </dl>
        <div class="hero-actions" aria-label="主要操作">
          <a class="btn btn--primary" href="#overview">查看项目概览</a>
          <a class="btn" href="#process">浏览实验过程</a>
          <a class="btn" href="../2026/07/10/软件开发工具实践实验报告_最终版/">查看完整报告</a>
          <a class="btn" href="https://thktyv8gjx-pixel.github.io/software-tools-practice-report/">访问公网站点</a>
          <a class="btn" href="https://github.com/thktyv8gjx-pixel/software-tools-practice-report">访问 GitHub 仓库</a>
        </div>
      </div>
      <aside class="hero-panel reveal" aria-label="项目事实摘要">
        <span class="panel-label">真实 Word 解析</span>
        <div class="hero-panel__stats">
          <div><strong data-count data-target="{stats['headings']}">0</strong><span>标题段落</span></div>
          <div><strong data-count data-target="{stats['tables']}">0</strong><span>表格</span></div>
          <div><strong data-count data-target="{stats['media']}">0</strong><span>媒体文件</span></div>
          <div><strong data-count data-target="{stats['captions']}">0</strong><span>图注</span></div>
        </div>
      </aside>
    </section>

    <section class="section" id="overview" aria-labelledby="overview-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Project Overview</p>
        <h2 id="overview-title">项目概览</h2>
        <p>本项目以“从本地开发到 Linux 服务器部署”为目标场景，把 Windows 开发、Hexo 静态构建、Git 版本管理、Ubuntu Server 运维、Nginx 发布和 GitHub Pages 公网部署串联为一条完整证据链。</p>
      </div>
      <div class="stats-grid">
        <article class="stat-card reveal"><strong data-count data-target="8">0</strong><span>项实验</span><p>按 Word 原顺序组织，从 Web 环境到 Web 软件部署。</p></article>
        <article class="stat-card reveal"><strong data-count data-target="6">0</strong><span>项考核核心</span><p>项目配置、Web 运行环境、Linux、账户、远程管理与部署。</p></article>
        <article class="stat-card reveal"><strong data-count data-target="4">0</strong><span>Ubuntu 节点</span><p>ubuntu-server、node1、node2、node3 组成多节点实验环境。</p></article>
        <article class="stat-card reveal"><strong data-count data-target="2">0</strong><span>发布通道</span><p>私网 Nginx 与公网 GitHub Pages 基于同一份报告内容。</p></article>
      </div>
      <div class="tech-cloud reveal" aria-label="技术栈">
        <span>Windows 11</span><span>Visual Studio Code</span><span>Markdown</span><span>Hexo</span><span>Git</span><span>GitHub</span><span>VirtualBox</span><span>Ubuntu Server</span><span>Linux</span><span>OpenSSH</span><span>SCP</span><span>rsync</span><span>Nginx</span><span>PowerShell</span><span>GitHub Actions</span><span>GitHub Pages</span>
      </div>
    </section>

    <section class="section" id="architecture" aria-labelledby="architecture-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Dual Channel Architecture</p>
        <h2 id="architecture-title">双通道技术架构</h2>
        <p>私网通道负责课程实验环境内的 Ubuntu Server 与 Nginx 发布，公网通道使用 GitHub main 分支、Actions、gh-pages 分支与 Pages 完成互联网访问。</p>
      </div>
      <div class="architecture-board reveal">
        <article class="pipeline pipeline--private">
          <h3>私网通道</h3>
          <ol>
            <li>Windows 11 开发主机</li>
            <li>VS Code / Markdown / Hexo</li>
            <li>Hexo 静态文件</li>
            <li>SSH / SCP / rsync</li>
            <li>Ubuntu Server</li>
            <li>Nginx</li>
            <li>192.168.56.101</li>
          </ol>
          <p>Ubuntu Server 固定 Host-Only 地址，Nginx 监听 80 端口，部署脚本以 HTTP 200 作为验收信号。</p>
        </article>
        <article class="pipeline pipeline--public">
          <h3>公网通道</h3>
          <ol>
            <li>Windows 11 开发主机</li>
            <li>Git 本地仓库</li>
            <li>GitHub main 分支</li>
            <li>GitHub Actions</li>
            <li>gh-pages 分支</li>
            <li>GitHub Pages</li>
            <li>公网访问</li>
          </ol>
          <p>main 保存源码，gh-pages 保存生成后的静态网站，Pages 使用 gh-pages 根目录作为发布来源。</p>
        </article>
      </div>
    </section>

    <section class="section" id="assessment" aria-labelledby="assessment-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Assessment Mapping</p>
        <h2 id="assessment-title">六项考核核心</h2>
        <p>以下六项来自 Word 第 4.3 节“实验成绩考核项目和内容”，每项都映射到实际实验章节、代表命令、关键配置和截图证据。</p>
      </div>
      <div class="assessment-grid">
        {assessment_cards()}
      </div>
    </section>

    <section class="section" id="process" aria-labelledby="process-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Experiment Process</p>
        <h2 id="process-title">实验过程</h2>
        <p>按照 Word 第 3 章原有顺序展示八项实验，每张卡片默认显示核心操作和结果，详细步骤可展开查看。</p>
      </div>
      <div class="experiment-list">
        {experiment_cards()}
      </div>
    </section>

    <section class="section" id="nodes" aria-labelledby="nodes-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Linux Cluster</p>
        <h2 id="nodes-title">Linux 多节点实验</h2>
        <p>在 ubuntu-server 基础上使用 VirtualBox 完整克隆 node1、node2、node3，并为每个节点配置独立主机名、machine-id、SSH host key 和 Host-Only 静态地址。</p>
      </div>
      <div class="topology reveal" role="img" aria-label="ubuntu-server、node1、node2、node3 多节点网络拓扑">
        <div class="topology__host">Windows Host-Only 适配器<br><strong>192.168.56.1/24</strong></div>
        <div class="topology__bus" aria-hidden="true"></div>
        {node_cards()}
      </div>
      <div class="node-notes reveal">
        <span>VirtualBox 完整克隆</span><span>NAT 网卡访问互联网</span><span>Host-Only 网卡固定管理入口</span><span>Netplan 静态 IP</span><span>/etc/hosts 节点解析</span><span>OpenSSH 服务</span><span>machine-id 重建</span><span>SSH host key 重建</span><span>节点互联验证</span>
      </div>
    </section>

    <section class="section" id="git" aria-labelledby="git-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Git & GitHub</p>
        <h2 id="git-title">Git 与 GitHub 发布链路</h2>
        <p>Word 中的 Git 实验覆盖 git init、status、add、commit、remote、push、origin、labserver、main、gh-pages、GitHub Actions 和 GitHub Pages。</p>
      </div>
      <div class="git-flow reveal">
        <span>main</span><span>GitHub Actions</span><span>gh-pages</span><span>pages build and deployment</span><span>GitHub Pages</span>
      </div>
      <div class="info-grid">
        <article class="info-card reveal"><h3>源码与远程</h3><p>GitHub 仓库为 thktyv8gjx-pixel/software-tools-practice-report；origin 指向 GitHub，labserver 指向 Ubuntu 裸 Git 仓库。</p></article>
        <article class="info-card reveal"><h3>Pages 最终设置</h3><p>Source：Deploy from a branch；Branch：gh-pages；Folder：/ (root)。图 53 作为最终设置截图。</p></article>
        <article class="info-card reveal"><h3>自动构建</h3><p>main 推送触发 Actions，执行 npm ci、Hexo 构建、public 产物验证，再发布到 gh-pages。</p></article>
      </div>
    </section>

    <section class="section" id="outcomes" aria-labelledby="outcomes-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Outcomes</p>
        <h2 id="outcomes-title">私网与公网成果</h2>
        <p>两个站点基于同一份实验报告内容：私网用于 Ubuntu Server/Nginx 实验验收，公网用于 GitHub Pages 成果展示。</p>
      </div>
      <div class="outcome-grid">
        <article class="outcome-card reveal">
          <span class="status-dot status-dot--private"></span>
          <h3>私网站点</h3>
          <p class="outcome-url">http://192.168.56.101</p>
          <ul>
            <li>Ubuntu Server 承载 Nginx 静态站点。</li>
            <li>SSH、SCP 与 rsync 完成构建产物同步。</li>
            <li>通过 curl 与浏览器验证 HTTP 200。</li>
            <li>仅在实验网络或对应主机环境可访问。</li>
          </ul>
          <a class="btn btn--primary" href="http://192.168.56.101">访问私网站点</a>
        </article>
        <article class="outcome-card reveal">
          <span class="status-dot status-dot--public"></span>
          <h3>公网站点</h3>
          <p class="outcome-url">https://thktyv8gjx-pixel.github.io/software-tools-practice-report/</p>
          <ul>
            <li>GitHub main 分支保存源码。</li>
            <li>GitHub Actions 构建并发布 gh-pages。</li>
            <li>GitHub Pages 提供 HTTPS 公网访问。</li>
            <li>图 55 显示公网完整样式访问成功。</li>
          </ul>
          <a class="btn btn--primary" href="https://thktyv8gjx-pixel.github.io/software-tools-practice-report/">访问公网站点</a>
        </article>
      </div>
    </section>

    <section class="section" id="issues" aria-labelledby="issues-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Troubleshooting</p>
        <h2 id="issues-title">典型问题与解决</h2>
        <p>问题卡片严格来自 Word 第 6 章，按“现象、错误截图、原因、排查、解决、成功截图、最终结果”组织。</p>
      </div>
      <div class="issue-grid">
        {issue_cards()}
      </div>
    </section>

    <section class="section" id="commands" aria-labelledby="commands-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Command Evidence</p>
        <h2 id="commands-title">关键命令</h2>
        <p>以下命令来自 Word 正文代码框，覆盖 Git、Hexo、Linux 用户权限、Netplan、SSH、SCP、rsync、Nginx、PowerShell 与 GitHub Actions。</p>
      </div>
      <div class="code-grid">
        {commands_html}
      </div>
    </section>

    <section class="section" id="gallery" aria-labelledby="gallery-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Screenshot Gallery</p>
        <h2 id="gallery-title">截图画廊</h2>
        <p>从 DOCX 内嵌媒体中精选 {len(selected_figures)} 张截图，覆盖工具链、Git、VirtualBox、多节点、SSH、rsync、Nginx、GitHub Actions、GitHub Pages 和典型错误。</p>
      </div>
      <div class="gallery-grid">
        {gallery_html}
      </div>
    </section>

    <section class="section" id="summary" aria-labelledby="summary-title">
      <div class="section-heading reveal">
        <p class="eyebrow">Reflection</p>
        <h2 id="summary-title">实验总结</h2>
        <p>Word 第 7 章强调，本次实践的价值不只是记住命令，而是把编辑器、构建工具、版本控制、虚拟机、Linux、网络、远程管理和 Web 服务器串联成完整工程链路。</p>
      </div>
      <div class="summary-grid">
        <article class="summary-card reveal"><h3>工具链综合理解</h3><p>Markdown 是内容源，Hexo 生成静态资源，Git 保存可追踪版本，SSH 承担安全传输，Linux 权限保护部署目录，Nginx 向浏览器提供服务。</p></article>
        <article class="summary-card reveal"><h3>工程化能力提升</h3><p>从 localhost 预览、GitHub 远程备份、Ubuntu 发布到自动部署脚本，项目成果具备可复现、可部署、可验证的工程价值。</p></article>
        <article class="summary-card reveal"><h3>故障定位方法</h3><p>先阅读错误信息和返回码，再判断问题属于路径、权限、服务、网络还是配置格式；修改后使用 netplan generate、sshd -t、nginx -t、curl 和浏览器复核。</p></article>
        <article class="summary-card reveal"><h3>安全与规范意识</h3><p>通过 webgroup、2770 权限、SSH 公钥认证和独立配置文件落实最小权限；对 rsync --delete 等自动化能力保持目录确认和结果检查。</p></article>
      </div>
    </section>
  </main>

  <footer class="footer">
    <div>
      <strong>软件开发工具实践实验报告</strong>
      <span>任垚橙 · 25051414</span>
      <span>杭州电子科技大学计算机学院</span>
    </div>
    <nav aria-label="页脚链接">
      <a href="#top">返回顶部</a>
      <a href="../2026/07/10/软件开发工具实践实验报告_最终版/">查看完整报告</a>
      <a href="https://github.com/thktyv8gjx-pixel/software-tools-practice-report">访问 GitHub 仓库</a>
      <a href="https://thktyv8gjx-pixel.github.io/software-tools-practice-report/">访问公网站点</a>
    </nav>
  </footer>

  <button class="back-to-top" type="button" aria-label="返回顶部">↑</button>

  <div class="lightbox" id="lightbox" role="dialog" aria-modal="true" aria-label="截图查看器" hidden>
    <div class="lightbox__backdrop" data-close></div>
    <div class="lightbox__panel" role="document">
      <button class="lightbox__close" type="button" aria-label="关闭截图查看器" data-close>×</button>
      <button class="lightbox__nav lightbox__nav--prev" type="button" aria-label="上一张截图" data-prev>‹</button>
      <img class="lightbox__image" src="" alt="">
      <button class="lightbox__nav lightbox__nav--next" type="button" aria-label="下一张截图" data-next>›</button>
      <p class="lightbox__caption" id="lightbox-caption"></p>
    </div>
  </div>

  <noscript>
    <p class="noscript">当前浏览器禁用了 JavaScript。正文、命令和截图仍可阅读，图片放大、复制按钮和导航高亮将不可用。</p>
  </noscript>
  <script src="./script.js" defer></script>
</body>
</html>
"""


def assessment_cards():
    cards = [
        ("项目与代码库配置", "项目配置、编辑器配置、Git 配置等。", "实验二：代码编辑与管理", "VS Code、Markdown、Git、GitHub", ".gitignore、Git 用户信息、origin 远程、main 分支", "git init / git status / git add / git commit / git remote / git push", "提交历史完整，GitHub 远程仓库保存源码。", "图10、图11、图12、图14"),
        ("Web 运行环境", "语言包、包管理器、关键软件包及配置等。", "实验一：Web 开发环境", "Node.js、npm、Hexo、VS Code、PowerShell", "npm registry、Hexo 项目结构、localhost:4000", "node -v / npm -v / npm install -g hexo-cli / hexo server", "Node/npm/Git/VS Code/Hexo 可用，站点可本地预览。", "图3、图4"),
        ("Linux 环境配置", "系统 profile 配置、用户配置、环境变量等。", "实验三、实验六", "VirtualBox、Ubuntu Server、Netplan、Bash", "NAT+Host-Only、静态 IP、.bashrc、/etc/profile.d", "netplan apply / hostnamectl / export / alias / source ~/.bashrc", "固定 IP 与环境变量持久生效，多节点独立身份配置完成。", "图17、图22、图34、图37"),
        ("Linux 账户配置", "账户配置、Home 配置、关键文件的权限配置等。", "实验五：Linux 账户与权限管理", "adduser、groupadd、chmod、chown", "webadmin、webgroup、/srv/software-tools-site、2770+setgid", "sudo adduser / sudo groupadd / sudo chmod 2770 / ls -ld", "共享目录组级协作可用，脚本执行权限验证成功。", "图30、图31、图33"),
        ("远程管理配置", "SSH 服务、端口配置、登录配置等。", "实验七：远程管理", "OpenSSH、SCP、rsync、裸 Git 仓库", "authorized_keys 权限、SSH Config、ubuntu-lab、labserver", "ssh / scp / rsync -av --delete / git push labserver main", "免密 SSH、跨节点传输、rsync 增量同步和裸 Git 仓库均成功。", "图38、图41、图42、图46"),
        ("Web 站点部署", "可运行的站点、基本功能、部署脚本。", "实验八：Web 软件部署", "Nginx、Hexo、PowerShell、GitHub Actions、GitHub Pages", "Nginx root、try_files、gh-pages、Pages root", "nginx -t / hexo generate / scp / rsync / curl -I", "私网 HTTP 200，公网 GitHub Pages 完整样式访问成功。", "图48、图50、图53、图54、图55"),
    ]
    return "\n".join(
        f"""
        <article class="assessment-card reveal">
          <h3>{e(title)}</h3>
          <dl>
            <div><dt>考核要求</dt><dd>{e(requirement)}</dd></div>
            <div><dt>对应实验</dt><dd>{e(experiment)}</dd></div>
            <div><dt>使用工具</dt><dd>{e(tools)}</dd></div>
            <div><dt>关键配置</dt><dd>{e(config)}</dd></div>
            <div><dt>代表命令</dt><dd>{e(command)}</dd></div>
            <div><dt>完成结果</dt><dd>{e(result)}</dd></div>
            <div><dt>对应截图</dt><dd>{e(figures)}</dd></div>
          </dl>
        </article>"""
        for title, requirement, experiment, tools, config, command, result, figures in cards
    )


def experiment_cards():
    experiments = [
        ("01", "实验一：Web开发环境", "安装并验证 Node.js、npm、Git、VS Code 与 Hexo，初始化 Hexo 项目并本地预览。", "PowerShell、Node.js、npm、Hexo、VS Code", "npm registry 设置、Hexo 目录结构、localhost:4000", "node -v；npm install -g hexo-cli；hexo server", "核心工具均可调用，Hexo 可从 Markdown 生成并预览静态网页。", "Web 运行环境", "图3、图4"),
        ("02", "实验二：代码编辑与管理", "使用 VS Code 编辑 Markdown，配置 Git 身份，完成首次提交、二次提交和 GitHub 远程推送。", "VS Code、Markdown、Git、GitHub、Git Graph", ".gitignore、origin、main、提交说明", "git status；git diff；git add；git commit；git push", "提交历史完整，源码同步到 GitHub 仓库。", "项目与代码库配置", "图10、图11、图12、图14"),
        ("03", "实验三：虚拟机安装与配置", "创建 Ubuntu Server 虚拟机，配置 NAT 与 Host-Only 双网卡、静态 IP、SSH，并扩展到三节点。", "VirtualBox、Ubuntu Server、Netplan、OpenSSH", "enp0s3 DHCP、enp0s8 静态地址、machine-id、SSH host key", "hostnamectl；netplan apply；ssh-keygen -A；systemctl enable --now ssh", "ubuntu-server 与 node1/node2/node3 地址固定且节点互联。", "Linux 环境配置", "图16、图17、图21、图25"),
        ("04", "实验四：Linux基本操作", "通过 SSH 使用系统信息、文件导航、APT 软件包管理、磁盘内存进程查看等基础命令。", "Linux Shell、apt、tree、ps、df、free", "实验目录范围、apt 索引、资源状态命令", "whoami；ip -br addr；apt update；tree；df -h；free -h", "能够在远程 Shell 中完成日常基础管理。", "Linux 环境配置", "图26、图27、图28、图29"),
        ("05", "实验五：Linux账户与权限管理", "创建 webadmin、webgroup 与共享网站目录，验证组权限、setgid 和脚本执行权限。", "adduser、groupadd、usermod、chmod、chown", "/srv/software-tools-site、root:webgroup、2770", "sudo adduser；sudo groupadd；sudo chmod 2770；chmod u+x", "组级协作有效，脚本权限从不可执行调整为可执行。", "Linux 账户配置", "图30、图31、图33"),
        ("06", "实验六：Linux环境配置", "配置 .bashrc、profile.d、PATH、环境变量、别名和自定义 labinfo 命令并验证持久性。", "Bash、profile.d、PATH、alias", "PROJECT_HOME、SITE_HOME、SOFTWARE_TOOLS_COURSE、~/bin", "export；alias；source ~/.bashrc；labinfo", "用户级和系统级环境配置在重新登录后仍然有效。", "Linux 环境配置", "图34、图35、图36、图37"),
        ("07", "实验七：远程管理", "完成 SCP 文件传输、ED25519 免密登录、SSH Config、Ubuntu 裸 Git 仓库、多节点 SCP 与 rsync。", "OpenSSH、SCP、rsync、Git", "authorized_keys 700/600、ubuntu-lab、labserver、--delete", "ssh；scp；rsync -av --delete；git push labserver main", "远程管理从手动输入 IP 和密码提升为可供脚本与 Git 复用的安全连接。", "远程管理配置", "图38、图41、图42、图46"),
        ("08", "实验八：Web软件部署", "安装 Nginx，配置静态站点，使用 Hexo 构建、SCP 上传、rsync 发布，并扩展到 GitHub Actions 与 Pages。", "Nginx、Hexo、PowerShell、GitHub Actions、GitHub Pages", "Nginx root、try_files、gh-pages、Pages root、部署脚本", "nginx -t；hexo generate；scp；rsync；curl -I", "私网 Nginx 返回 200，公网 GitHub Pages 完整样式访问成功。", "Web 站点部署", "图47、图48、图50、图53、图54、图55"),
    ]
    return "\n".join(
        f"""
        <article class="experiment-card reveal">
          <div class="experiment-card__head">
            <span>{num}</span>
            <div><h3>{e(title)}</h3><p>{e(purpose)}</p></div>
          </div>
          <dl class="experiment-brief">
            <div><dt>核心操作</dt><dd>{e(config)}</dd></div>
            <div><dt>最终结果</dt><dd>{e(result)}</dd></div>
            <div><dt>对应考核</dt><dd>{e(assessment)}</dd></div>
          </dl>
          <details>
            <summary>查看步骤、工具和截图证据</summary>
            <dl>
              <div><dt>使用工具</dt><dd>{e(tools)}</dd></div>
              <div><dt>关键命令</dt><dd>{e(commands)}</dd></div>
              <div><dt>关键配置</dt><dd>{e(config)}</dd></div>
              <div><dt>典型截图</dt><dd>{e(figures)}</dd></div>
            </dl>
          </details>
        </article>"""
        for num, title, purpose, tools, config, commands, result, assessment, figures in experiments
    )


def node_cards():
    nodes = [
        ("ubuntu-server", "192.168.56.101", "Nginx / Git / SSH 主服务器"),
        ("node1", "192.168.56.102", "多节点 SSH、SCP、rsync 发起节点"),
        ("node2", "192.168.56.103", "跨节点传输与同步目标节点"),
        ("node3", "192.168.56.104", "互联验证节点"),
    ]
    return "\n".join(
        f"""<article class="node-card"><h3>{e(name)}</h3><strong>{e(ip)}</strong><p>{e(role)}</p></article>"""
        for name, ip, role in nodes
    )


def issue_cards():
    issues = [
        ("6.1 Node.js命令无法识别", "PowerShell 执行 node -v 提示 CommandNotFoundException。", "图56", "Node.js 尚未安装或安装目录未加入 PATH。", "确认命令不可用、重新安装并保留 Add to PATH，关闭旧终端后重开。", "安装 Node.js 后重新验证 node -v。", "图3", "node -v 正常输出 v24.18.0。"),
        ("6.2 PowerShell阻止npm.ps1", "node 可用但 npm -v 被执行策略阻止。", "图57", "PowerShell 执行策略阻止 npm.ps1。", "用 npm.cmd -v 判断 npm 本体是否存在，再调整当前用户执行策略。", "执行 Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned。", "图3", "npm 命令恢复，后续可安装 Hexo CLI。"),
        ("6.3 localhost页面拒绝连接", "浏览器显示 ERR_CONNECTION_REFUSED。", "图58", "地址写成 localhost4000，且 Hexo 服务进程已退出。", "核对 URL 格式和 hexo server 进程状态。", "改为 http://localhost:4000 并保持服务运行。", "图5、图6", "本地页面可正常显示。"),
        ("6.4 虚拟机无法读取启动介质", "VirtualBox 提示 Could not read from the boot medium。", "图59", "Ubuntu ISO 未正确挂载。", "检查虚拟机存储控制器和启动顺序。", "挂载 ISO 到虚拟光驱并设置光驱优先启动。", "图18", "Ubuntu Server 能正常启动安装流程。"),
        ("6.5 Ubuntu安装在GRUB阶段失败", "安装在 curtin in-target/grub-pc 阶段失败。", "图60", "虚拟机配置和安装介质组合不稳定。", "删除失败虚拟机，重新创建动态 VDI，关闭 EFI 与 3D 加速。", "使用完整 Ubuntu Server 22.04.5 LTS 镜像并避免在线升级。", "图18、图19", "第二次安装成功并识别网络接口。"),
        ("6.6 Netplan YAML解析失败", "netplan generate 提示 Invalid YAML: aliases are not supported。", "图61", "误输入字符和缩进不规范。", "清空配置，使用空格重新输入并先做语法检查。", "执行 chmod 600、netplan generate 和 netplan apply。", "图19、图22", "固定 IP 生效。"),
        ("6.7 SSH快捷主机名无法解析", "ssh ubuntu-lab 提示 Could not resolve hostname。", "图62", "配置可能保存为 config.txt 或缺少 Host 行。", "检查 SSH Config 文件名、编码和 ssh -G 输出。", "用 ASCII 编码重新写入 HostName、User、Port 和 IdentityFile。", "图41", "ssh ubuntu-lab 可直接连接。"),
        ("6.8 Hexo误加载PowerShell部署脚本", "Hexo 将 deploy.ps1 当作 JavaScript 扩展加载并报 SyntaxError。", "图63", "脚本被放入 Hexo 具有特殊含义的 scripts 目录。", "核对 Hexo 加载规则和报错文件路径。", "将脚本移到 tools 目录重新运行。", "图52", "部署脚本返回 HTTP 200 并完成推送。"),
        ("6.9 VirtualBox克隆时磁盘空间不足", "完整克隆 node2、node3 出现 VERR_DISK_FULL。", "图64", "宿主机剩余空间不足，完整克隆会复制整套虚拟磁盘。", "删除失败节点和无效 VDI，释放空间后逐台克隆并检查磁盘路径与容量。", "清理磁盘空间后重新完整克隆。", "图21", "三个 Ubuntu 节点完整克隆成功。"),
        ("6.10 克隆节点启动卡住与SSH超时", "node1 停留在 Loading essential drivers，ping 与 22 端口测试超时。", "图65", "克隆节点继承模板身份与旧网络状态。", "按“虚拟机启动、网卡地址、SSH 服务”分层排查。", "重新生成 machine-id 和 SSH host key，确认 Host-Only 地址后重启。", "图22、图25", "节点恢复启动，SSH 与互联验证成功。"),
        ("6.11 GitHub Pages主题CSS返回404", "Pages 首页可打开但只显示默认排版，/css/style.css 返回 404。", "图66", "仓库名曾拼写为 sofware-tools-practice-report，与正确资源根路径不一致。", "核对 gh-pages 文件、Raw 文件、仓库真实名称、Pages 路径和资源 URL。", "将仓库重命名为正确拼写，更新 origin 并重新触发部署。", "图53、图54、图55", "CSS 请求恢复 200，公网完整主题正常显示。"),
    ]
    return "\n".join(
        f"""
        <article class="issue-card reveal">
          <h3>{e(title)}</h3>
          <dl>
            <div><dt>问题现象</dt><dd>{e(symptom)}</dd></div>
            <div><dt>错误截图</dt><dd>{e(error_fig)}</dd></div>
            <div><dt>原因分析</dt><dd>{e(reason)}</dd></div>
            <div><dt>排查过程</dt><dd>{e(diagnosis)}</dd></div>
            <div><dt>解决操作</dt><dd>{e(solution)}</dd></div>
            <div><dt>成功截图</dt><dd>{e(success_fig)}</dd></div>
            <div><dt>最终结果</dt><dd>{e(result)}</dd></div>
          </dl>
        </article>"""
        for title, symptom, error_fig, reason, diagnosis, solution, success_fig, result in issues
    )


STYLE_CSS = r"""
:root {
  color-scheme: dark;
  --bg: #050914;
  --bg-2: #081424;
  --panel: rgba(12, 24, 42, 0.82);
  --panel-strong: rgba(16, 34, 58, 0.94);
  --line: rgba(126, 209, 255, 0.2);
  --line-strong: rgba(126, 209, 255, 0.42);
  --text: #e7f2ff;
  --muted: #a8bdd4;
  --soft: #6f88a5;
  --cyan: #5bdcff;
  --blue: #4a8dff;
  --green: #61e6b2;
  --amber: #ffd166;
  --danger: #ff7b91;
  --shadow: 0 20px 70px rgba(0, 0, 0, 0.38);
  --radius: 8px;
  --nav-h: 72px;
  --max: 1180px;
  --mono: Consolas, "SFMono-Regular", "Liberation Mono", Menlo, monospace;
  --sans: "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", "Segoe UI", Arial, sans-serif;
}

* {
  box-sizing: border-box;
}

html {
  scroll-behavior: smooth;
  scroll-padding-top: calc(var(--nav-h) + 24px);
}

body {
  margin: 0;
  overflow-x: hidden;
  background:
    linear-gradient(180deg, rgba(5, 9, 20, 0.88), rgba(5, 9, 20, 0.98)),
    radial-gradient(circle at 20% 0%, rgba(38, 126, 255, 0.18), transparent 32rem),
    radial-gradient(circle at 80% 10%, rgba(91, 220, 255, 0.12), transparent 30rem),
    repeating-linear-gradient(90deg, rgba(91, 220, 255, 0.035) 0 1px, transparent 1px 80px),
    repeating-linear-gradient(0deg, rgba(91, 220, 255, 0.028) 0 1px, transparent 1px 80px),
    var(--bg);
  color: var(--text);
  font-family: var(--sans);
  line-height: 1.65;
}

body::before {
  content: "";
  position: fixed;
  inset: 0;
  pointer-events: none;
  background-image:
    radial-gradient(circle, rgba(231, 242, 255, 0.7) 0 1px, transparent 1.5px),
    radial-gradient(circle, rgba(91, 220, 255, 0.55) 0 1px, transparent 1.5px);
  background-size: 120px 120px, 190px 190px;
  background-position: 10px 20px, 70px 90px;
  opacity: 0.18;
}

img {
  display: block;
  max-width: 100%;
  height: auto;
}

a {
  color: inherit;
}

button,
a {
  -webkit-tap-highlight-color: transparent;
}

:focus-visible {
  outline: 3px solid var(--cyan);
  outline-offset: 3px;
}

.skip-link {
  position: fixed;
  top: 12px;
  left: 12px;
  z-index: 999;
  transform: translateY(-140%);
  padding: 10px 14px;
  border-radius: var(--radius);
  background: var(--cyan);
  color: #001321;
  font-weight: 700;
}

.skip-link:focus {
  transform: translateY(0);
}

.site-header {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  z-index: 100;
  min-height: var(--nav-h);
  border-bottom: 1px solid var(--line);
  background: rgba(5, 9, 20, 0.82);
  backdrop-filter: blur(16px);
}

.nav {
  max-width: var(--max);
  margin: 0 auto;
  min-height: var(--nav-h);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 24px;
  padding: 0 24px;
}

.nav__brand {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  min-width: 220px;
  text-decoration: none;
  font-weight: 800;
  letter-spacing: 0;
}

.brand-mark {
  width: 28px;
  aspect-ratio: 1;
  border-radius: 50%;
  border: 2px solid var(--cyan);
  box-shadow: inset 0 0 16px rgba(91, 220, 255, 0.45), 0 0 18px rgba(91, 220, 255, 0.28);
}

.nav__menu {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  flex-wrap: wrap;
  gap: 4px;
  padding: 0;
  margin: 0;
  list-style: none;
}

.nav__menu a {
  display: inline-flex;
  align-items: center;
  min-height: 40px;
  padding: 8px 10px;
  border-radius: var(--radius);
  color: var(--muted);
  text-decoration: none;
  font-size: 14px;
  font-weight: 700;
}

.nav__menu a:hover,
.nav__menu a.is-active {
  color: var(--text);
  background: rgba(91, 220, 255, 0.12);
}

.nav__toggle {
  display: none;
  width: 44px;
  height: 44px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.04);
  color: var(--text);
}

.nav__toggle span {
  display: block;
  width: 20px;
  height: 2px;
  margin: 5px auto;
  background: currentColor;
}

.section {
  position: relative;
  max-width: var(--max);
  margin: 0 auto;
  padding: 96px 24px;
  scroll-margin-top: calc(var(--nav-h) + 28px);
}

.hero {
  min-height: 100svh;
  display: grid;
  grid-template-columns: minmax(0, 1.45fr) minmax(300px, 0.55fr);
  align-items: center;
  gap: 34px;
  padding-top: calc(var(--nav-h) + 56px);
  padding-bottom: 64px;
}

.hero__grid {
  position: absolute;
  inset: var(--nav-h) 24px 24px;
  border: 1px solid rgba(91, 220, 255, 0.12);
  border-radius: var(--radius);
  background:
    linear-gradient(115deg, rgba(91, 220, 255, 0.08), transparent 38%),
    linear-gradient(300deg, rgba(74, 141, 255, 0.12), transparent 42%);
  opacity: 0.72;
}

.hero__content,
.hero-panel {
  position: relative;
  z-index: 1;
}

.eyebrow {
  margin: 0 0 14px;
  color: var(--cyan);
  font-size: 13px;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.12em;
}

h1,
h2,
h3 {
  margin: 0;
  line-height: 1.18;
  letter-spacing: 0;
}

h1 {
  max-width: 860px;
  font-size: clamp(42px, 8vw, 86px);
}

h2 {
  font-size: clamp(30px, 4.8vw, 50px);
}

h3 {
  font-size: 20px;
}

.hero__subtitle {
  max-width: 790px;
  margin: 22px 0 28px;
  color: var(--muted);
  font-size: clamp(18px, 2.4vw, 26px);
}

.hero-meta {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
  margin: 0 0 28px;
}

.hero-meta div,
.hero-panel,
.stat-card,
.assessment-card,
.experiment-card,
.issue-card,
.code-card,
.gallery-card,
.summary-card,
.info-card,
.outcome-card {
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: var(--panel);
  box-shadow: var(--shadow);
}

.hero-meta div {
  padding: 12px 14px;
}

dt {
  color: var(--soft);
  font-size: 13px;
}

dd {
  margin: 0;
}

.hero-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
}

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 44px;
  padding: 10px 16px;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.045);
  color: var(--text);
  text-decoration: none;
  font-weight: 800;
}

.btn:hover,
.btn--primary {
  border-color: transparent;
  background: linear-gradient(135deg, var(--blue), var(--cyan));
  color: #001321;
}

.hero-panel {
  padding: 22px;
}

.panel-label {
  color: var(--green);
  font-weight: 800;
}

.hero-panel__stats {
  display: grid;
  gap: 16px;
  margin-top: 18px;
}

.hero-panel__stats div {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 14px;
  padding-bottom: 12px;
  border-bottom: 1px solid var(--line);
}

.hero-panel__stats strong,
.stat-card strong {
  color: var(--cyan);
  font-family: var(--mono);
  font-size: 38px;
}

.section-heading {
  max-width: 850px;
  margin-bottom: 32px;
}

.section-heading p:last-child {
  margin: 14px 0 0;
  color: var(--muted);
  font-size: 17px;
}

.stats-grid,
.assessment-grid,
.issue-grid,
.summary-grid,
.outcome-grid,
.info-grid {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 16px;
}

.stat-card {
  min-height: 210px;
  padding: 22px;
}

.stat-card span {
  display: block;
  margin-bottom: 10px;
  font-weight: 800;
}

.stat-card p,
.assessment-card dd,
.experiment-card p,
.issue-card dd,
.summary-card p,
.info-card p,
.outcome-card li {
  color: var(--muted);
}

.tech-cloud,
.node-notes {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 20px;
}

.tech-cloud span,
.node-notes span {
  padding: 8px 12px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(91, 220, 255, 0.08);
  color: #cfeaff;
  font-weight: 700;
}

.architecture-board {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 18px;
}

.pipeline {
  padding: 22px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: var(--panel-strong);
}

.pipeline h3 {
  margin-bottom: 18px;
}

.pipeline ol {
  display: grid;
  gap: 12px;
  padding: 0;
  margin: 0 0 18px;
  list-style: none;
}

.pipeline li {
  position: relative;
  padding: 12px 14px 12px 40px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.035);
  font-weight: 800;
}

.pipeline li::before {
  content: "";
  position: absolute;
  left: 15px;
  top: 50%;
  width: 10px;
  height: 10px;
  border-radius: 50%;
  transform: translateY(-50%);
  background: var(--cyan);
}

.pipeline p {
  margin: 0;
  color: var(--muted);
}

.assessment-grid {
  grid-template-columns: repeat(3, minmax(0, 1fr));
}

.assessment-card,
.experiment-card,
.issue-card,
.summary-card,
.info-card,
.outcome-card {
  padding: 20px;
}

.assessment-card dl,
.experiment-card dl,
.issue-card dl {
  display: grid;
  gap: 12px;
  margin: 18px 0 0;
}

.assessment-card dt,
.experiment-card dt,
.issue-card dt {
  color: var(--cyan);
  font-weight: 800;
}

.experiment-list {
  display: grid;
  gap: 16px;
}

.experiment-card__head {
  display: grid;
  grid-template-columns: 58px minmax(0, 1fr);
  gap: 16px;
  align-items: start;
}

.experiment-card__head > span {
  display: grid;
  place-items: center;
  width: 58px;
  aspect-ratio: 1;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  color: var(--cyan);
  font-family: var(--mono);
  font-weight: 900;
  font-size: 22px;
}

.experiment-brief {
  grid-template-columns: repeat(3, minmax(0, 1fr));
  margin: 16px 0;
}

details {
  border-top: 1px solid var(--line);
  padding-top: 14px;
}

summary {
  cursor: pointer;
  color: var(--cyan);
  font-weight: 800;
}

.topology {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 16px;
  align-items: stretch;
  padding: 22px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(8, 20, 36, 0.78);
}

.topology__host {
  grid-column: 1 / -1;
  padding: 16px;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  text-align: center;
  background: rgba(91, 220, 255, 0.08);
}

.topology__bus {
  grid-column: 1 / -1;
  height: 2px;
  background: linear-gradient(90deg, transparent, var(--cyan), transparent);
}

.node-card {
  min-height: 170px;
  padding: 18px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.04);
}

.node-card strong {
  display: block;
  margin: 12px 0;
  color: var(--green);
  font-family: var(--mono);
}

.git-flow {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 12px;
}

.git-flow span {
  position: relative;
  padding: 16px;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: var(--panel);
  text-align: center;
  font-weight: 900;
}

.git-flow span:not(:last-child)::after {
  content: "→";
  position: absolute;
  right: -14px;
  top: 50%;
  transform: translateY(-50%);
  color: var(--cyan);
}

.info-grid,
.outcome-grid {
  grid-template-columns: repeat(2, minmax(0, 1fr));
  margin-top: 18px;
}

.outcome-url {
  overflow-wrap: anywhere;
  color: var(--cyan);
  font-family: var(--mono);
}

.status-dot {
  display: inline-block;
  width: 12px;
  aspect-ratio: 1;
  margin-bottom: 12px;
  border-radius: 50%;
}

.status-dot--private {
  background: var(--amber);
}

.status-dot--public {
  background: var(--green);
}

.issue-grid {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}

.code-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
}

.code-card {
  overflow: hidden;
}

.code-card__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  border-bottom: 1px solid var(--line);
  background: rgba(91, 220, 255, 0.08);
  font-weight: 900;
}

.copy-btn {
  min-width: 58px;
  min-height: 34px;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.06);
  color: var(--text);
  font-weight: 800;
  cursor: pointer;
}

.copy-btn.is-copied {
  border-color: var(--green);
  color: var(--green);
}

.copy-btn.is-failed {
  border-color: var(--danger);
  color: var(--danger);
}

.code-block {
  position: relative;
  margin: 0;
  max-width: 100%;
  overflow-x: auto;
  padding: 44px 16px 16px;
  background: #050b12;
  color: #dff6ff;
  font-family: var(--mono);
  font-size: 14px;
  line-height: 1.6;
}

.code-block::before {
  content: attr(data-language);
  position: absolute;
  top: 12px;
  left: 16px;
  color: var(--soft);
  font-size: 12px;
  font-weight: 800;
}

.gallery-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
}

.gallery-card {
  overflow: hidden;
}

.gallery-card__media {
  width: 100%;
  aspect-ratio: 16 / 10;
  padding: 0;
  border: 0;
  background: #020712;
  cursor: zoom-in;
}

.gallery-card__media img {
  width: 100%;
  height: 100%;
  object-fit: contain;
}

.gallery-card figcaption {
  display: grid;
  gap: 6px;
  padding: 14px;
}

.gallery-card strong {
  line-height: 1.35;
}

.gallery-card span,
.gallery-card small {
  color: var(--muted);
}

.summary-grid {
  grid-template-columns: repeat(4, minmax(0, 1fr));
}

.footer {
  display: flex;
  justify-content: space-between;
  gap: 24px;
  max-width: var(--max);
  margin: 0 auto;
  padding: 34px 24px 48px;
  border-top: 1px solid var(--line);
}

.footer div,
.footer nav {
  display: flex;
  flex-wrap: wrap;
  gap: 12px 18px;
  align-items: center;
}

.footer span {
  color: var(--muted);
}

.footer a {
  color: var(--cyan);
  text-decoration: none;
  font-weight: 800;
}

.back-to-top {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 80;
  width: 46px;
  height: 46px;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  background: var(--panel-strong);
  color: var(--cyan);
  font-size: 22px;
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.2s ease, transform 0.2s ease;
}

.back-to-top.is-visible {
  opacity: 1;
  pointer-events: auto;
}

.lightbox[hidden] {
  display: none;
}

.lightbox {
  position: fixed;
  inset: 0;
  z-index: 200;
  display: grid;
  place-items: center;
  padding: 20px;
}

.lightbox__backdrop {
  position: absolute;
  inset: 0;
  background: rgba(0, 0, 0, 0.78);
}

.lightbox__panel {
  position: relative;
  z-index: 1;
  width: min(1180px, 100%);
  max-height: 92vh;
  display: grid;
  grid-template-columns: 50px minmax(0, 1fr) 50px;
  grid-template-rows: minmax(0, 1fr) auto;
  gap: 10px;
  padding: 14px;
  border: 1px solid var(--line-strong);
  border-radius: var(--radius);
  background: #050914;
}

.lightbox__image {
  grid-column: 2;
  max-height: 76vh;
  width: 100%;
  object-fit: contain;
}

.lightbox__caption {
  grid-column: 1 / -1;
  margin: 0;
  color: var(--muted);
  text-align: center;
}

.lightbox__close,
.lightbox__nav {
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: rgba(255, 255, 255, 0.06);
  color: var(--text);
  cursor: pointer;
}

.lightbox__close {
  position: absolute;
  top: 12px;
  right: 12px;
  width: 42px;
  height: 42px;
  font-size: 26px;
}

.lightbox__nav {
  align-self: center;
  height: 72px;
  font-size: 38px;
}

.lightbox__nav--prev {
  grid-column: 1;
}

.lightbox__nav--next {
  grid-column: 3;
}

.noscript {
  max-width: var(--max);
  margin: 0 auto 24px;
  padding: 12px 24px;
  color: var(--amber);
}

.reveal {
  opacity: 0;
  transform: translateY(18px);
  transition: opacity 0.55s ease, transform 0.55s ease;
}

.reveal.is-visible {
  opacity: 1;
  transform: translateY(0);
}

@media (max-width: 1120px) {
  .hero,
  .architecture-board,
  .outcome-grid,
  .info-grid {
    grid-template-columns: 1fr;
  }

  .stats-grid,
  .summary-grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .assessment-grid,
  .gallery-grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .git-flow {
    grid-template-columns: 1fr;
  }

  .git-flow span:not(:last-child)::after {
    content: "↓";
    right: 50%;
    top: auto;
    bottom: -22px;
    transform: translateX(50%);
  }
}

@media (max-width: 880px) {
  .nav {
    padding: 0 16px;
  }

  .nav__toggle {
    display: block;
  }

  .nav__menu {
    position: absolute;
    top: var(--nav-h);
    left: 16px;
    right: 16px;
    display: none;
    padding: 12px;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    background: rgba(5, 9, 20, 0.96);
  }

  .nav__menu.is-open {
    display: grid;
    grid-template-columns: 1fr;
  }

  .nav__menu a {
    width: 100%;
  }

  .section {
    padding: 76px 16px;
  }

  .hero {
    min-height: auto;
    padding-top: calc(var(--nav-h) + 44px);
  }

  .hero-meta,
  .stats-grid,
  .assessment-grid,
  .issue-grid,
  .code-grid,
  .gallery-grid,
  .summary-grid,
  .topology,
  .experiment-brief {
    grid-template-columns: 1fr;
  }

  .hero-actions .btn {
    width: 100%;
  }

  .experiment-card__head {
    grid-template-columns: 1fr;
  }

  .lightbox__panel {
    grid-template-columns: 42px minmax(0, 1fr) 42px;
    padding: 10px;
  }
}

@media (max-width: 520px) {
  h1 {
    font-size: 38px;
  }

  h2 {
    font-size: 30px;
  }

  .nav__brand {
    min-width: 0;
    font-size: 14px;
  }

  .hero-panel__stats strong,
  .stat-card strong {
    font-size: 30px;
  }

  .footer {
    display: grid;
  }
}

@media (prefers-reduced-motion: reduce) {
  *,
  *::before,
  *::after {
    scroll-behavior: auto !important;
    transition-duration: 0.001ms !important;
    animation-duration: 0.001ms !important;
    animation-iteration-count: 1 !important;
  }

  .reveal {
    opacity: 1;
    transform: none;
  }
}
"""


SCRIPT_JS = r"""
(function () {
  "use strict";

  document.documentElement.classList.add("js");

  var navMenu = document.getElementById("nav-menu");
  var navToggle = document.querySelector(".nav__toggle");
  var navLinks = Array.prototype.slice.call(document.querySelectorAll(".nav__menu a[href^='#']"));
  var sections = navLinks
    .map(function (link) {
      var target = document.querySelector(link.getAttribute("href"));
      return target ? { link: link, section: target } : null;
    })
    .filter(Boolean);

  if (navToggle && navMenu) {
    navToggle.addEventListener("click", function () {
      var open = navMenu.classList.toggle("is-open");
      navToggle.setAttribute("aria-expanded", String(open));
    });
    navLinks.forEach(function (link) {
      link.addEventListener("click", function () {
        navMenu.classList.remove("is-open");
        navToggle.setAttribute("aria-expanded", "false");
      });
    });
  }

  if ("IntersectionObserver" in window) {
    var navObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) {
            return;
          }
          sections.forEach(function (item) {
            item.link.classList.toggle("is-active", item.section === entry.target);
          });
        });
      },
      { rootMargin: "-35% 0px -55% 0px", threshold: 0.01 }
    );
    sections.forEach(function (item) {
      navObserver.observe(item.section);
    });
  } else if (sections.length) {
    sections[0].link.classList.add("is-active");
  }

  var revealItems = Array.prototype.slice.call(document.querySelectorAll(".reveal"));
  if ("IntersectionObserver" in window) {
    var revealObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
            revealObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12 }
    );
    revealItems.forEach(function (item) {
      revealObserver.observe(item);
    });
  } else {
    revealItems.forEach(function (item) {
      item.classList.add("is-visible");
    });
  }

  function animateNumber(element) {
    var target = Number(element.getAttribute("data-target") || "0");
    if (!Number.isFinite(target)) {
      element.textContent = "0";
      return;
    }
    var start = 0;
    var duration = 900;
    var startTime = null;
    function tick(time) {
      if (startTime === null) {
        startTime = time;
      }
      var progress = Math.min((time - startTime) / duration, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      element.textContent = String(Math.round(start + (target - start) * eased));
      if (progress < 1) {
        window.requestAnimationFrame(tick);
      }
    }
    window.requestAnimationFrame(tick);
  }

  var counters = Array.prototype.slice.call(document.querySelectorAll("[data-count]"));
  if ("IntersectionObserver" in window) {
    var counterObserver = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            animateNumber(entry.target);
            counterObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.4 }
    );
    counters.forEach(function (counter) {
      counterObserver.observe(counter);
    });
  } else {
    counters.forEach(animateNumber);
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(function (resolve, reject) {
      var textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "");
      textarea.style.position = "fixed";
      textarea.style.top = "-999px";
      document.body.appendChild(textarea);
      textarea.select();
      try {
        var ok = document.execCommand("copy");
        document.body.removeChild(textarea);
        ok ? resolve() : reject(new Error("copy command failed"));
      } catch (error) {
        document.body.removeChild(textarea);
        reject(error);
      }
    });
  }

  Array.prototype.slice.call(document.querySelectorAll("[data-copy]")).forEach(function (button) {
    button.addEventListener("click", function () {
      var card = button.closest(".code-card");
      var code = card ? card.querySelector("code") : null;
      if (!code) {
        return;
      }
      copyText(code.textContent)
        .then(function () {
          button.textContent = "已复制";
          button.classList.add("is-copied");
          button.classList.remove("is-failed");
          window.setTimeout(function () {
            button.textContent = "复制";
            button.classList.remove("is-copied");
          }, 1400);
        })
        .catch(function () {
          button.textContent = "请手动复制";
          button.classList.add("is-failed");
          button.classList.remove("is-copied");
          window.setTimeout(function () {
            button.textContent = "复制";
            button.classList.remove("is-failed");
          }, 1800);
        });
    });
  });

  var lightbox = document.getElementById("lightbox");
  var lightboxImage = lightbox ? lightbox.querySelector(".lightbox__image") : null;
  var lightboxCaption = lightbox ? lightbox.querySelector(".lightbox__caption") : null;
  var triggers = Array.prototype.slice.call(document.querySelectorAll(".lightbox-trigger"));
  var activeIndex = 0;
  var lastFocus = null;

  function figureCaption(trigger) {
    var figure = trigger.closest("figure");
    var title = figure ? figure.querySelector("figcaption strong") : null;
    var text = figure ? figure.querySelector("figcaption span") : null;
    var meta = figure ? figure.querySelector("figcaption small") : null;
    return [title, text, meta]
      .filter(Boolean)
      .map(function (item) {
        return item.textContent.trim();
      })
      .join(" · ");
  }

  function showImage(index) {
    if (!lightbox || !lightboxImage || !lightboxCaption || !triggers.length) {
      return;
    }
    activeIndex = (index + triggers.length) % triggers.length;
    var trigger = triggers[activeIndex];
    var image = trigger.querySelector("img");
    if (!image) {
      return;
    }
    lightboxImage.src = image.currentSrc || image.src;
    lightboxImage.alt = image.alt;
    lightboxCaption.textContent = figureCaption(trigger);
  }

  function openLightbox(index) {
    if (!lightbox) {
      return;
    }
    lastFocus = document.activeElement;
    lightbox.hidden = false;
    document.body.style.overflow = "hidden";
    showImage(index);
    var close = lightbox.querySelector("[data-close]");
    if (close) {
      close.focus();
    }
  }

  function closeLightbox() {
    if (!lightbox || lightbox.hidden) {
      return;
    }
    lightbox.hidden = true;
    document.body.style.overflow = "";
    if (lightboxImage) {
      lightboxImage.src = "";
    }
    if (lastFocus && typeof lastFocus.focus === "function") {
      lastFocus.focus();
    }
  }

  triggers.forEach(function (trigger, index) {
    trigger.addEventListener("click", function () {
      openLightbox(index);
    });
  });

  if (lightbox) {
    Array.prototype.slice.call(lightbox.querySelectorAll("[data-close]")).forEach(function (button) {
      button.addEventListener("click", closeLightbox);
    });
    var prev = lightbox.querySelector("[data-prev]");
    var next = lightbox.querySelector("[data-next]");
    if (prev) {
      prev.addEventListener("click", function () {
        showImage(activeIndex - 1);
      });
    }
    if (next) {
      next.addEventListener("click", function () {
        showImage(activeIndex + 1);
      });
    }
  }

  document.addEventListener("keydown", function (event) {
    if (!lightbox || lightbox.hidden) {
      return;
    }
    if (event.key === "Escape") {
      closeLightbox();
    } else if (event.key === "ArrowLeft") {
      showImage(activeIndex - 1);
    } else if (event.key === "ArrowRight") {
      showImage(activeIndex + 1);
    }
  });

  var topButton = document.querySelector(".back-to-top");
  if (topButton) {
    topButton.addEventListener("click", function () {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    window.addEventListener(
      "scroll",
      function () {
        topButton.classList.toggle("is-visible", window.scrollY > 520);
      },
      { passive: true }
    );
  }
})();
"""


def write_files(stats):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "index.html").write_text(build_html(stats), encoding="utf-8", newline="\n")
    (OUT_DIR / "style.css").write_text(STYLE_CSS.strip() + "\n", encoding="utf-8", newline="\n")
    (OUT_DIR / "script.js").write_text(SCRIPT_JS.strip() + "\n", encoding="utf-8", newline="\n")


def main():
    stats = read_docx_metadata()
    copy_selected_images()
    write_files(stats)
    print(
        "generated report-demo: "
        f"paragraphs={stats['paragraphs']} headings={stats['headings']} "
        f"top_headings={stats['top_headings']} tables={stats['tables']} "
        f"media={stats['media']} image_refs={stats['image_refs']} captions={stats['captions']} "
        f"selected_images={len(SELECTED_IMAGES)}"
    )


if __name__ == "__main__":
    main()
