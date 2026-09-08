# PDF / DOCX 转结构化 Markdown

这个自包含 Python 工具把 PDF 和 Word DOCX 转成结构化 Markdown，提取图片、表格、代码、标题和列表。原生 PDF 在本地解析；扫描件或缺少可用文字层的 PDF 可以自动转到 MinerU OCR。

## 快速开始

需要 Python 3.10 或更新版本。克隆仓库后，Windows 可直接运行：

```powershell
git clone https://github.com/leentau/pdf2md.git
cd pdf2md
.\install.bat
.\pdf2md.bat "C:\文档\manual.pdf" --overwrite --report
```

`install.bat` 会在项目目录创建隔离的 `.venv` 并安装 `requirements.txt` 中的全部依赖。也可以跳过安装命令，第一次运行 `pdf2md.bat` 时会自动安装。

PowerShell 用户可以直接使用 `.\install.ps1` 和 `.\pdf2md.ps1`。Linux/macOS 使用：

```bash
git clone https://github.com/leentau/pdf2md.git
cd pdf2md
chmod +x install.sh pdf2md
./install.sh
./pdf2md /path/to/manual.pdf --overwrite --report
```

所有 Python 包都安装在项目自己的 `.venv` 中，不修改系统 Python 环境。依赖中的 `pypandoc_binary` 自带 Pandoc；若系统已安装 Pandoc，程序会优先使用系统版本。

## 项目结构

```text
pdf2md/
├─ native_pdf_to_md.py       # 统一命令入口、原生 PDF 解析和自动分流
├─ mineru_pdf_to_md.py       # MinerU API、分片、续传和结果合并
├─ word_to_md.py             # DOCX 结构预处理与 Pandoc 转换
├─ pandoc_word_filter.lua    # 合并相邻 Word 代码段
├─ requirements.txt          # 运行依赖
├─ CHANGELOG.md              # 每次公开更新的修改与验证记录
├─ install.ps1 / install.sh  # 创建 .venv 并安装、检查依赖
├─ pdf2md.bat / pdf2md       # 自动安装并运行的入口
└─ test_native_pdf_to_md.py  # 自动化回归测试
```

源代码、安装器、启动器和测试都位于同一个目录；不依赖仓库外部的 Python 文件。每次公开更新都会同时更新 `CHANGELOG.md`，记录修改原因、行为变化、验证结果和兼容性影响。`.venv`、Token、转换缓存和报告不会上传到 GitHub。

## 输出目录规则

每个源文件都会得到一个独立文件夹。单个文件和文件夹输入采用不同的输出规则。

指定单个 `a\manual.pdf` 或 `a\manual.docx` 时，在源文件同目录建立 `manual` 文件夹：

```text
a/
├─ manual.pdf（或 manual.docx）
└─ manual/
   ├─ manual.md
   └─ images/
      ├─ page-014-image-01.png
      └─ ...
```

指定文件夹 `a` 时，在 `a` 的同级创建 `a_md`，递归转换并镜像 `a` 的目录结构。PDF 和 Word 文档可以混合放置。

程序首先判断 PDF 是否包含足够的可见原生文字。纯原生文字 PDF 在本地处理，不上传；整页扫描图、隐藏 OCR 文字层或原生文字不足的 PDF 会自动调用 MinerU 精准解析 API，并强制开启 OCR。原生解析会读取：

- PDF 书签与原生字体层级，转换成严格对应的 Markdown 标题层级；
- 原生文字和坐标，去除重复页眉、页脚并恢复段落；位于页面顶部但不重复的栏目标题会保留；
- 项目符号、编号步骤、Note/Tip/Warning 等提示块；
- Courier 等等宽字体代码块；正文、项目符号和命令交替出现在同一 PDF 文字块时会按行拆分，避免整块误判为代码；
- 相邻或跨页连续的代码片段自动合并为一个代码块，并保留原始缩进；
- PDF 矢量表格线和单元格，转换成 Markdown 表格；
- PDF 嵌入的原始图片，保存到 `images` 并写入相对引用。
- 自动识别常见双栏论文和横向三栏 Quick Reference 版式，按从左到右、每栏从上到下的阅读顺序输出，并保留跨栏页标题；
- 彩色横条分区标题会与其下方的小节标题、正文分开；无框的“命令/说明”双列清单会恢复为 Markdown 表格；
- 带 `Fig.` 标题的纯矢量论文插图会渲染到 `images`，不使用 OCR。
- IEEE 论文的罗马数字章节和字母小节会映射为对应 Markdown 层级；
- 对带合并表头的有框表格，会依据原生矢量表格线补全漏行和列标题。
- Computer Modern 数学字体会组合为 LaTeX 上下标、分式、求和式和根式；
- 保留论文正文的首行缩进，并将整栏参考文献按 `[序号]` 逐条分段。
- 图片引用使用不含引文方括号的短标签（如 `![Fig. 1](images/...)`），避免 Markdown 预览器截断链接。
- 连续字体 span 会先合并；整段 Medium/Bold 作为区域基准样式，不生成成排的 `**` 标记；
- `{user1, user2}@domain` 形式的论文合并邮箱会展开为两个普通邮箱地址。

Word 文档转换会保留：

- Word `Heading 1` 到 `Heading 6` 与 Markdown `#` 到 `######` 的严格对应关系，并在写入前核验数量；
- 自动编号、项目列表、超链接和内部书签；
- 简单 Markdown 表格以及带合并单元格的 HTML 表格；
- 内嵌和浮动图片，统一保存到同级 `images` 文件夹；
- Courier、Consolas 等等宽字体段落的缩进，并把相邻代码段合并成一个围栏代码块；
- 由分页布局拆开的连续 `Body Text` 句子会按结构合并；包含 `-option`/`--option` 的通用命令语法会避免被斜体和公式渲染器拆成逐字竖排；
- 接受修订后的可见正文，并且不会执行文档宏。

当前只正式支持 DOCX；旧 `.doc` 及其他 Office 格式等有实际样品后再增加。

## 配置 MinerU OCR

不要把 Token 写入 Python 源码或命令行。推荐设置环境变量：

```powershell
$env:MINERU_TOKEN = "你的 MinerU Token"
```

也可以复制 `.env.example` 为项目目录下的 `.env`，然后填写 Token：

```dotenv
MINERU_TOKEN=你的_MinerU_Token
```

程序会自动读取程序目录、程序上一级、当前目录或输入目录中的 `.env`。系统环境变量的优先级高于 `.env`；`.env` 已由 `.gitignore` 排除，不会上传到 GitHub。

也可以把 Token 单独保存为 `mineru_token.txt`。程序会依次在程序目录、程序目录的上一级和输入 PDF 目录查找；或者显式指定：

```powershell
.\pdf2md.bat .\paper `
  --mineru-token-file .\mineru_token.txt
```

自动分流规则：

1. `--route auto` 是默认值，单文件和目录批处理都不必额外指定。目录模式会对每个 PDF 分别判断。
2. 分流颗粒度固定为整份 PDF，只有本地原生解析和 MinerU OCR 两种结果，不会在一份 Markdown 中混合两种引擎。
3. PDF 具有足够的可见原生文字、没有检测到整页扫描页，并且没有超复杂矢量页面时，整份 PDF 使用本地原生元素解析。
4. 任意页面检测为整页扫描图、只有 `ignore-text` 隐藏文字层，或整份 PDF 的原生可见文字总量不足时，整份 PDF 使用 MinerU OCR。
5. 任意页面自身及其嵌套 Form XObject 的压缩绘图内容合计超过 2 MiB 时，为避免本地展开数十万个矢量路径而卡住，整份 PDF 自动使用 MinerU OCR。
6. 显式指定 `--route native` 会禁止自动上传；遇到扫描页或超复杂矢量页时明确报错，由用户决定是否改用 MinerU。
7. MinerU 使用 `vlm` 模型，并开启 OCR、表格和公式识别；可用 `--mineru-model pipeline` 切换模型。
8. 超过 API 单文件限制的文档会按 180 页或 190 MiB 的安全阈值自动拆分，最后合并 Markdown 和图片。

也可以强制指定路线，便于对同一份大文件分别验证两种处理方式：

```powershell
# 把 ptug.pdf 当作纯原生文字 PDF：本地逐页处理，不上传
.\pdf2md.bat .\paper\goo\ptug.pdf `
  --route native --report --overwrite

# 把同一个 ptug.pdf 当作扫描图片 PDF：强制 MinerU OCR、自动分片、断点续传和合并
.\pdf2md.bat .\paper\goo\ptug.pdf `
  --route mineru-ocr --report --overwrite
```

MinerU 分片阈值可以调整，但不能超过官方单任务限制：

```powershell
--max-pages-per-chunk 180 --max-api-file-mb 190
```

大文件 OCR 的进度保存在输出根目录的 `.mineru_cache`。网络中断或程序退出后，再运行同一条命令会从已上传/已完成的分片继续；结果 ZIP 下载中断时也会通过 HTTP Range 从 `.part` 文件继续。上传和下载过程每增加约 8 MiB 显示一次进度，连续 180 秒没有响应才会超时。每个 MinerU 签名上传任务只执行一次 PUT；失败后由外层任务刷新状态、申请新的上传地址并重试，避免状态长期停在 `uploading`。全部完成后缓存自动删除。分片合并不会插入源 PDF 中不存在的 Markdown 分隔线，所有图片会重命名后汇总到最终 `images` 文件夹。

MinerU OCR 会把需要识别的 PDF 上传到 MinerU 服务。敏感文档使用前应确认允许发送到外部服务。若要完全禁止上传，使用 `--disable-mineru-ocr`。

若 Clash 等代理为 `cdn-mineru.openxlab.org.cn` 返回 `198.18.x.x` fake-IP 并导致 TLS EOF，下载器会通过公共 DNS 获取当前 CDN 地址后直连。直连仍使用原域名进行 SNI 和证书校验，不会关闭 HTTPS 验证。也可用 `MINERU_CDN_IP` 指定可信的公开 IPv4 地址。

## 转换单个 PDF

```powershell
.\pdf2md.bat .\a\tshell_lbist_user.pdf
```

输出固定为 PDF 同目录的 `a/tshell_lbist_user/tshell_lbist_user.md`，图片位于 `a/tshell_lbist_user/images/`。

覆盖已有结果并输出结构核验报告（报告写入输出根目录的 `_reports`，不会改变单个 PDF 文件夹结构）：

```powershell
.\pdf2md.bat .\a\tshell_lbist_user.pdf --overwrite --report
```

## 转换 Word 文档

```powershell
.\pdf2md.bat .\tshell_ijtag_user.docx `
  --overwrite --report
```

单文件输出为 `tshell_ijtag_user/tshell_ijtag_user.md` 和 `tshell_ijtag_user/images/`。当前支持 `.docx`；`--route` 和 MinerU 参数只作用于 PDF。

## 批量转换

```powershell
.\pdf2md.bat .\pdf目录
```

目录模式自动递归遍历所有子目录中的 PDF 和 Word 文档，不再要求 `--recursive`。程序在 `pdf目录` 的同级建立 `pdf目录_md`，并完整保留源目录结构：

```text
父目录/
├─ pdf目录/
│  ├─ top.pdf
│  └─ manuals/
│     └─ chip.pdf
└─ pdf目录_md/
   ├─ top/
   │  ├─ top.md
   │  └─ images/
   └─ manuals/
      └─ chip/
         ├─ chip.md
         └─ images/
```

目录批处理再次运行时，默认跳过已经存在对应 Markdown 的文件，只重试上次失败或尚未处理的文件。需要重新生成全部结果时才使用 `--overwrite`。

输出位置是固定规则：单文件放在 PDF 旁边；文件夹批处理放在同级 `<输入目录名>_md`。旧命令中的 `-o/--output-dir` 参数仍可被解析，但会被忽略并显示提示。

## 标题层级规则

1. PDF 书签是主依据。书签 1 到 6 级严格对应 Markdown `#` 到 `######`。
2. 标题只在书签指定页的正文区域匹配，目录页和重复页眉不会误判。
3. PDF 中没有书签记录的可见标题，才使用原生字号、粗体和颜色补充判断。
4. 使用 `--report` 后，输出根目录的 `_reports/<PDF文件名>.json` 会列出未匹配书签，便于人工核验。

## 限制

- 未配置 MinerU Token 时，扫描版、隐藏 OCR 文字层或文字不足的 PDF 会停止并提示如何配置，不会静默使用质量较差的隐藏文字。
- MinerU 精准解析 API 当前限制单个任务不超过 200 MB、200 页；程序会在此限制之前自动拆分。
- Markdown 不支持跨页合并单元格。跨页表格会按 PDF 页分别输出。
- 没有图题的纯矢量示意图可能无法自动确定边界；嵌入图像会原样保存，带 `Fig.` 图题的矢量图会按原页面区域渲染保存。
- 单页自身及其嵌套 Form XObject 的压缩绘图内容合计超过 2 MiB 时，默认的 `--route auto` 会把整份 PDF 交给 MinerU OCR；这会将文档上传到 MinerU。使用 `--route native` 或 `--disable-mineru-ocr` 可以禁止上传，但程序会停止并说明触发自动分流的页码和大小。
- Indexed、ICCBased 等特殊 PDF 图片会先规范化为 PNG 兼容色彩空间；soft mask 分辨率与彩图不同时会按彩图尺寸缩放，仍无法解码的单图会使用页面区域渲染兜底，不会中断整份文档。

## 测试

安装完成后运行：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -p "test_*.py" -v
```

测试使用运行时生成的小型 PDF、DOCX 和模拟 MinerU 响应，不会上传文档。

## 安全说明

- 不要把真实 Token 写入源码、README、`.env.example` 或命令行历史。
- `.gitignore` 已排除 `.env`、`mineru_token.txt`、虚拟环境和运行缓存。
- MinerU OCR 会把需要识别的 PDF 上传到外部服务；敏感文件请使用 `--route native` 或 `--disable-mineru-ocr`。
