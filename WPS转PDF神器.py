"""WPS Batch PDF Converter / WPS Word 批量转 PDF。

在 Windows 上调用本机 WPS Office，递归地将 DOC、DOCX 和 WPS 文档导出为 PDF。
项目主页：https://github.com/jedliuai/wps-batch-pdf-converter
"""

import os
import sys
import time
import tkinter as tk
from tkinter import filedialog
import win32com.client

__version__ = "1.0.1"

def select_folder():
    """弹出窗口选择文件夹"""
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口，只显示选择框
    folder_path = filedialog.askdirectory(title="请选择包含Word文档的文件夹")
    return folder_path

def main():
    print("************************************************")
    print("         WPS Word 批量转 PDF 小工具")
    print("  适用于 WPS 可正常打开文档的绿盾环境")
    print("************************************************\n")

    print("正在等待选择文件夹...")
    folder = select_folder()

    if not folder:
        print("❌ 未选择文件夹，程序即将退出。")
        time.sleep(2)
        return

    print(f"📂 已选择文件夹: {folder}")
    print("正在扫描文件，请稍候...\n")

    # 1. 扫描所有文件
    tasks = []
    for root, dirs, files in os.walk(folder):
        for f in files:
            # 兼容 doc, docx 和 wps 格式
            if f.lower().endswith(('.doc', '.docx', '.wps')) and not f.startswith('~$'):
                full_path = os.path.join(root, f)
                pdf_path = os.path.join(root, os.path.splitext(f)[0] + ".pdf")
                # 如果PDF不存在才添加任务
                if not os.path.exists(pdf_path):
                    tasks.append((full_path, pdf_path))

    total = len(tasks)
    if total == 0:
        print("🎉 该文件夹下没有需要转换的文件（或者都已经转好了）。")
        input("\n按【回车键】退出...")
        return

    print(f"📊 发现 {total} 个文件需要转换，准备启动 WPS...")

    # 2. 启动 WPS
    wps = None
    try:
        # 尝试两种常见的WPS注册名
        try:
            wps = win32com.client.Dispatch("Kwps.Application")
        except:
            wps = win32com.client.Dispatch("Wps.Application")
            
        wps.Visible = True  # 保持可见，以兼容部分绿盾环境的文档访问策略
        wps.DisplayAlerts = False
    except Exception as e:
        print("\n❌ 启动 WPS 失败！请确认电脑上安装了 WPS Office。")
        print(f"错误详情: {e}")
        input("\n按【回车键】退出...")
        return

    # 3. 开始转换循环
    print("🚀 开始转换...\n")
    success_count = 0
    fail_count = 0

    for index, (doc_path, pdf_path) in enumerate(tasks, 1):
        filename = os.path.basename(doc_path)
        # 打印进度条
        print(f"[{index}/{total}] 正在处理: {filename}")
        
        try:
            doc = wps.Documents.Open(doc_path)
            doc.SaveAs2(pdf_path, FileFormat=17) # 17 代表 PDF
            doc.Close(SaveChanges=False)
            success_count += 1
        except Exception as e:
            print(f"    ⚠️ 转换失败: {filename}")
            print(f"    错误信息: {e}")
            fail_count += 1
            # 尝试关闭可能卡住的文档
            try: doc.Close(SaveChanges=False)
            except: pass

    # 4. 收尾
    wps.Quit()
    
    print("\n" + "="*40)
    print(f"✅ 全部完成！")
    print(f"成功: {success_count} 个")
    print(f"失败: {fail_count} 个")
    print("="*40)
    
    # 这里的 input 是为了防止窗口处理完瞬间消失
    input("\n按【回车键】退出程序...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # 全局错误捕获，防止闪退
        print(f"程序发生严重错误: {e}")
        input("按【回车键】退出...")
