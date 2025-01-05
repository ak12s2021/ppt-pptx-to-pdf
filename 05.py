import os
import win32com.client
import pythoncom


def ppt_to_pdf(ppt_file, pdf_file):
    try:
        # 初始化COM线程
        pythoncom.CoInitialize()

        # 使用win32com
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = True  # 保持窗口可见

        # 打开演示文稿
        presentation = powerpoint.Presentations.Open(ppt_file, WithWindow=True)

        # 另存为PDF
        presentation.SaveAs(pdf_file, FileFormat=32)  # 32 是 ppSaveAsPDF

        # 关闭演示文稿
        presentation.Close()

        print(f"成功转换: {ppt_file}")

    except Exception as e:
        print(f"转换失败 {ppt_file}: {e}")

    finally:
        # 退出PowerPoint
        powerpoint.Quit()
        # 释放COM资源
        pythoncom.CoUninitialize()


def convert_all_ppts(folder_path):
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return

    # 计数器
    converted_count = 0
    total_count = 0

    for filename in os.listdir(folder_path):
        # 支持 .ppt 和 .pptx 文件
        if filename.lower().endswith((".pptx", ".ppt")):
            total_count += 1
            ppt_file = os.path.join(folder_path, filename)
            pdf_file = os.path.splitext(ppt_file)[0] + ".pdf"

            # 避免重复转换
            if not os.path.exists(pdf_file):
                ppt_to_pdf(ppt_file, pdf_file)
                converted_count += 1

    print(f"总文件数: {total_count}")
    print(f"成功转换: {converted_count}")


if __name__ == "__main__":
    folder_path = input("请输入包含PPT文件的文件夹路径: ")
    convert_all_ppts(folder_path)