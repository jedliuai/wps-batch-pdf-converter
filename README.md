# WPS Batch PDF Converter｜WPS Word 批量转 PDF

[![Latest release](https://img.shields.io/github/v/release/jedliuai/wps-batch-pdf-converter?label=Windows%20版下载)](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest)
[![Total downloads](https://img.shields.io/github/downloads/jedliuai/wps-batch-pdf-converter/total?label=downloads)](https://github.com/jedliuai/wps-batch-pdf-converter/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue.svg)](#运行条件与边界)

这是一个专门面向 **绿盾等透明加密办公环境**的 Windows + WPS Office 批量 Word 转 PDF 工具：普通 Python/headless 转换程序无法正常读取受保护文档，但当前用户有权限通过 WPS 打开时，本工具让获授权的 WPS 逐个完成 PDF 导出。

> **使用前提：你必须能够在这台电脑上用 WPS 正常打开目标文档。** 如果 WPS 自己也打不开，或当前账号没有读取、导出权限，本工具同样无法处理；它不负责解密，也不会绕过企业安全策略。

之所以选择 WPS，是因为在这个受限场景里，它既能打开被允许访问的加密文档，单文件导出 PDF 又**速度快、版式稳定**；作者实际使用的 Adobe 转换链路虽然也能处理，但速度太慢。WPS 官方提供的单文件导出 PDF 可以免费使用，原生“批量 PDF 导出”则属于会员功能（不同版本和地区的权益可能变化，以 WPS 当前说明为准）。本工具不破解 WPS、不修改会员状态，而是把正常的单文件导出自动串成批处理。[查看 WPS 官方说明](https://www.wps.com/academy/how-to-convert-word-excel-ppt-to-pdf-for-free-in-wps-office-quick-tutorials-1863083/)

选择一个文件夹后，程序会递归扫描全部子文件夹，将 `.doc`、`.docx`、`.wps` 文档导出为同目录 PDF；已有同名 PDF 会被跳过，不上传文件，也不覆盖现有结果。

> [下载最新版 Windows EXE](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest) · [查看源码运行方式](#从源码运行) · [反馈问题](https://github.com/jedliuai/wps-batch-pdf-converter/issues)

## 先判断它是否适合你

| 你的情况 | 是否适合 |
| --- | --- |
| 文档受绿盾等软件保护，普通脚本读不到，但 WPS 可以正常打开和导出 | **适合，这正是本项目的核心场景** |
| Adobe 等获授权软件也能处理，但在实际环境中转换太慢 | 适合，前提是 WPS 转换更快且能正常导出 |
| 文档没有加密，也没有软件访问限制 | 通常没必要，优先考虑更通用的 Python、LibreOffice 或 headless 批量方案 |
| WPS 无法打开文档，或当前用户没有读取/导出权限 | **不适合，本工具不能解密或绕过权限** |
| 需要在 macOS、Linux、服务器或无人值守环境运行 | 不适合，当前实现依赖 Windows 桌面版 WPS |

## 30 秒开始使用

1. 确认电脑已安装桌面版 **WPS Office**，并且你有权正常打开需要转换的文档。
2. 从 [Releases](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest) 下载 `WPS-Batch-PDF-Converter-*-Windows-x64.exe`，双击运行。
3. 在弹出的窗口中选择文档文件夹；程序会递归处理子文件夹，并在每个源文件旁生成同名 PDF。

转换时 WPS 窗口会保持可见，这是兼容部分文档保护环境所需的正常行为。首次运行未签名 EXE 时，Windows 可能显示 SmartScreen 提示；请只从本仓库 Release 下载，并可用 Release 说明中的 SHA-256 校验文件。

## 它解决什么问题

在普通未加密环境中，用户可以选择许多成熟的通用转换方案；真正困难的是透明加密后，Python 进程可能读不到文档，而获授权的 WPS 或 Adobe 可以。本项目把重复操作交给能够正常访问文档的本机 WPS，一次解决“受限环境下只能走获授权软件”“Adobe 链路太慢”“WPS 原生批量功能可能需要会员”三个痛点：

```text
选择根文件夹
├─ 客户 A/报价单.docx  → 客户 A/报价单.pdf
├─ 客户 A/说明书.wps   → 客户 A/说明书.pdf
├─ 客户 B/合同.doc      → 客户 B/合同.pdf
└─ 客户 B/已归档.docx   → 已有同名 PDF，自动跳过
```

典型场景包括：

- 一次性归档多层目录中的 Word/WPS 文件；
- 批量生成给客户、同事或系统上传用的 PDF；
- 在只能通过本机 WPS 正常读取文档的办公环境中减少重复操作；
- 中断后再次运行，只补做尚未生成 PDF 的文件。

## 真实能力与默认行为

| 项目 | 行为 |
| --- | --- |
| 输入格式 | `.doc`、`.docx`、`.wps` |
| 扫描范围 | 所选文件夹及全部子文件夹 |
| 输出位置 | 与源文档相同目录、相同文件名 |
| 已有 PDF | 自动跳过，不覆盖 |
| 临时文件 | 自动忽略以 `~$` 开头的文件 |
| 转换引擎 | 本机已安装的 WPS Office，通过 Windows COM 调用 |
| 文件传输 | 完全本地处理，不上传到云端 |
| 过程反馈 | 控制台显示总数、当前进度、成功数与失败数 |

## 运行条件与边界

- 仅支持 **Windows**；需要已安装可正常使用的桌面版 WPS Office。
- 本工具调用 WPS 自己的 PDF 导出能力，最终版式以 WPS 打开文档时的效果为准。
- 对天锐绿盾等透明加密环境的兼容方式，是让获授权的 WPS 保持可见并走正常打开、导出流程。只有 WPS 本身能正常打开并获准导出的文档才可能转换成功。程序**不会自行解密、绕过权限或规避企业安全策略**。
- 同目录已有同名 PDF 时会直接跳过。如需重新生成，请先自行备份或移走旧 PDF。
- 当前采用控制台显示进度；大量或复杂文档的速度取决于 WPS 和电脑性能。

### 这是在绕过 WPS 会员吗？

不是。本工具不会解锁或调用 WPS 的会员批量转换入口，也不会修改 WPS、账号或授权状态。它只对每个有权正常打开的文档，依次调用 WPS 桌面版自身的单文件 PDF 导出能力。请遵守 WPS 的许可条款和所在组织的安全政策。

## 从源码运行

需要 Python 3.9+（64 位 Windows 推荐）：

```powershell
git clone https://github.com/jedliuai/wps-batch-pdf-converter.git
cd wps-batch-pdf-converter
python -m pip install -r requirements.txt
python ".\WPS转PDF神器.py"
```

自行构建单文件 EXE：

```powershell
python -m pip install pyinstaller
pyinstaller --clean ".\WPS转PDF神器.spec"
```

## 隐私与安全

程序只遍历你主动选择的本地文件夹，并通过本机 WPS 打开和导出文档；源码中没有网络上传逻辑。公开 Release 暂未做商业代码签名，因此对安全敏感的环境，建议直接审阅源码后运行或自行构建。

报告问题时，请提供 Windows 版本、WPS 版本、输入扩展名、是否处于文档保护环境，以及控制台错误信息。请勿上传含敏感内容的原始文档。

## 许可证与作者

本项目采用 [MIT License](LICENSE)。如果它确实替你省下了重复劳动，欢迎 Star、分享给同样需要的人，或提交 Issue 帮助它覆盖更多真实场景。

由 [Jed (@jedliuai)](https://github.com/jedliuai) 维护。更多面向真实工作流的开源自动化工具，可在作者的 [GitHub 主页](https://github.com/jedliuai?tab=repositories) 查看。
