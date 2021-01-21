# -*- coding: UTF-8 -*-
"""
Auto-merge tool for duplex scanning
Author: SKR
Version: 2.2
"""

import PyPDF2
import os
import time
import win32api
import win32con


pdf1_path = ".\\1.pdf"
pdf2_path = ".\\2.pdf"

current_time = time.strftime("%Y%m%d_%H%M%S", time.localtime(time.time()))
output_name = f"pdf_merged_at_{current_time}"
output_path = f".\\{output_name}.pdf"

try:
    pdf_1 = PyPDF2.PdfFileReader(open(pdf1_path, 'rb'), strict=False)
except FileNotFoundError:
    win32api.MessageBox(0, "文档 <1.pdf> 不存在", "任务中止", win32con.MB_ICONWARNING)
    exit()
except PyPDF2.utils.PdfReadError:
    win32api.MessageBox(0, "无法解析 <1.pdf> 请检查文档完整性", "任务中止", win32con.MB_ICONWARNING)
    exit()
except:
    win32api.MessageBox(0, "在解析 <1.pdf> 时遇到未知错误", "任务中止", win32con.MB_ICONWARNING)
    exit()

try:
    pdf_2 = PyPDF2.PdfFileReader(open(pdf2_path, 'rb'), strict=False)
except FileNotFoundError:
    win32api.MessageBox(0, "文档 <2.pdf> 不存在", "任务中止", win32con.MB_ICONWARNING)
    exit()
except PyPDF2.utils.PdfReadError:
    win32api.MessageBox(0, "无法解析 <2.pdf> 请检查文档完整性", "任务中止", win32con.MB_ICONWARNING)
    exit()
except:
    win32api.MessageBox(0, "在解析 <2.pdf> 时遇到未知错误", "任务中止", win32con.MB_ICONWARNING)
    exit()

n_page_pdf1 = pdf_1.getNumPages()
n_page_pdf2 = pdf_2.getNumPages()
if n_page_pdf1 != n_page_pdf2:
    win32api.MessageBox(0, f"待合并文档页数不相等\nPDF1: {n_page_pdf1} | PDF2: {n_page_pdf2}", "任务中止", win32con.MB_ICONWARNING)
    exit()

pdf_merged = PyPDF2.PdfFileWriter()
n = n_page_pdf1
for i in range(n):
    pdf_merged.addPage(pdf_1.getPage(i))
    pdf_merged.addPage(pdf_2.getPage(n - 1 - i))

try:
    pdf_merged.write(open(output_path, 'wb'))
except:
    win32api.MessageBox(0, "保存合并文档时遇到错误\n请检查权限或重试", "任务中止", win32con.MB_ICONWARNING)
    exit()
else:
    result = win32api.MessageBox(0, f"合并文档已生成于\n{output_name}\n是否打开？", "任务完成", win32con.MB_YESNO)
    if result == win32con.IDYES:
        os.system("start " + output_path)
