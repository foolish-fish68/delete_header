import os
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog

# 定义Word常量
wdStory = 6
wdCharacter = 1
wdParagraph = 1  # 段落单位
wdMove = 0
wdExtend = 1


def select_word_files():
    """弹出文件选择对话框，允许选择多个Word文件（无数量限制）"""
    # 创建隐藏的Tkinter主窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口，只显示文件选择对话框

    # 弹出文件选择对话框，支持多选
    file_paths = filedialog.askopenfilenames(
        title="选择Word文件",
        filetypes=[("Word Files", "*.doc;*.docx"), ("All Files", "*.*")]
    )

    # 处理选中的文件路径（过滤非Word文件）
    valid_files = []
    for path in file_paths:
        if os.path.isfile(path) and path.lower().endswith(('.doc', '.docx')):
            valid_files.append(path)

    return valid_files


def process_word_file(file_path):
    """处理单个Word文件，删除【试题区】及其之前的所有内容，并清除开头空行"""
    try:
        # 初始化COM组件
        pythoncom.CoInitialize()

        # 初始化Word应用程序
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 后台运行，不显示界面

        # 打开文档
        doc = word.Documents.Open(os.path.abspath(file_path))

        # 查找"【试题区】"
        find = doc.Content.Find
        find.Text = "【试题区】"
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = True

        # 执行查找
        if find.Execute():
            # 获取查找到的范围
            found_range = doc.Content.Duplicate
            found_range.SetRange(find.Parent.Start, find.Parent.End)

            # 选中从文档开始到找到的文本（包括"【试题区】"）
            doc.Range(0, found_range.End).Select()

            # 删除选中的内容
            word.Selection.Delete()

            # 删除开头可能存在的空行
            # 移动到文档开头
            word.Selection.HomeKey(Unit=wdStory)

            # 循环检查并删除所有开头的空段落
            while True:
                # 检查当前段落是否为空（只包含段落标记）
                if word.Selection.Paragraphs(1).Range.Text.strip() == "":
                    # 删除当前空段落
                    word.Selection.Paragraphs(1).Range.Delete()
                    # 再次移动到开头，准备检查下一个可能的空段落
                    word.Selection.HomeKey(Unit=wdStory)
                else:
                    # 找到非空段落，退出循环
                    break

            # 保存并关闭文档
            doc.Save()
            doc.Close()
            print(f"已成功处理: {file_path}")
        else:
            doc.Close(SaveChanges=0)  # 不保存关闭
            print(f"未在文件中找到【试题区】: {file_path}")

    except Exception as e:
        print(f"处理文件时出错 {file_path}: {str(e)}")
    finally:
        # 确保Word进程关闭
        if 'word' in locals():
            word.Quit()
        # 释放COM组件
        pythoncom.CoUninitialize()


def main():
    print("请选择要处理的Word文件...")
    word_files = select_word_files()

    if not word_files:
        print("未选择任何文件，程序退出。")
        return

    print(f"共选择了 {len(word_files)} 个文件，开始处理...")
    for file in word_files:
        process_word_file(file)

    print("所有文件处理完毕！")


if __name__ == "__main__":
    main()