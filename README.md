# WPS Batch PDF Converter｜WPS Word 批量转 PDF

[![Latest release](https://img.shields.io/github/v/release/jedliuai/wps-batch-pdf-converter?label=Windows%20版下载)](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest)
[![Total downloads](https://img.shields.io/github/downloads/jedliuai/wps-batch-pdf-converter/total?label=downloads)](https://github.com/jedliuai/wps-batch-pdf-converter/releases)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue.svg)](#运行条件与边界)

一个面向 Windows + WPS Office 的本地批量 Word 转 PDF 工具。选择一个文件夹后，它会自动扫描全部子文件夹，将 `.doc`、`.docx`、`.wps` 文档导出为同目录 PDF；已有同名 PDF 会被跳过，不上传文件，也不覆盖现有结果。

特别适合归档资料、批量交付文件，以及需要让 WPS 保持可见才能正常导出的部分天锐绿盾（Tipray）办公环境。

> [下载最新版 Windows EXE](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest) · [查看源码运行方式](#从源码运行) · [反馈问题](https://github.com/jedliuai/wps-batch-pdf-converter/issues)

## 30 秒开始使用

1. 确认电脑已安装桌面版 **WPS Office**，并且你有权正常打开需要转换的文档。
2. 从 [Releases](https://github.com/jedliuai/wps-batch-pdf-converter/releases/latest) 下载 `WPS-Batch-PDF-Converter-*-Windows-x64.exe`，双击运行。
3. 在弹出的窗口中选择文档文件夹；程序会递归处理子文件夹，并在每个源文件旁生成同名 PDF。

转换时 WPS 窗口会保持可见，这是兼容部分文档保护环境所需的正常行为。首次运行未签名 EXE 时，Windows 可能显示 SmartScreen 提示；请只从本仓库 Release 下载，并可用 Release 说明中的 SHA-256 校验文件。

## 它解决什么问题

手工逐个打开文档、另存为 PDF，在文件多、目录深或文档分散时非常耗时。这个工具把重复操作交给本机 WPS：

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
- 对天锐绿盾环境的兼容方式是让 WPS 保持可见并走正常打开、导出流程。它**不会解密文档、绕过权限或规避企业安全策略**；没有读取权限的文件仍无法转换。
- 同目录已有同名 PDF 时会直接跳过。如需重新生成，请先自行备份或移走旧 PDF。
- 当前采用控制台显示进度；大量或复杂文档的速度取决于 WPS 和电脑性能。

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
