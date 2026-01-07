import os
import sys
from docx2pdf import convert
from pypdf import PdfWriter


def merge_pdfs():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(base_dir, "output")
    merge_file = os.path.join(base_dir, "merge.pdf")

    if not os.path.exists(output_dir):
        print("错误: 输出目录不存在，无法合并。")
        return

    # 获取所有 pdf 文件并按文件名排序
    pdf_files = sorted([f for f in os.listdir(output_dir) if f.endswith(".pdf")])

    if not pdf_files:
        print("没有找到需要合并的 PDF 文件。")
        return

    print(f"\n开始合并 PDF...")
    merger = PdfWriter()

    for filename in pdf_files:
        path = os.path.join(output_dir, filename)
        print(f"正在合并: {filename}")
        merger.append(path)

    try:
        merger.write(merge_file)
        merger.close()
        print(f"\n合并完成！已保存为: {merge_file}")
    except Exception as e:
        print(f"合并 PDF 时出错: {e}")


def batch_convert():
    # 获取当前脚本所在目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input")
    output_dir = os.path.join(base_dir, "output")

    # 检查输入目录是否存在
    if not os.path.exists(input_dir):
        print(f"错误: 输入目录 '{input_dir}' 不存在。")
        return

    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"已创建输出目录: {output_dir}")

    print(f"开始转换...")
    print(f"输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")

    # 获取所有 docx 文件
    files = [
        f
        for f in os.listdir(input_dir)
        if f.endswith(".docx") and not f.startswith("~$")
    ]

    if not files:
        print("没有找到需要转换的 .docx 文件。")
    else:
        converted_count = 0
        skipped_count = 0

        for filename in files:
            input_path = os.path.join(input_dir, filename)
            # 生成目标 PDF 文件名（替换扩展名）
            pdf_filename = os.path.splitext(filename)[0] + ".pdf"
            output_path = os.path.join(output_dir, pdf_filename)

            # 检查目标文件是否已存在
            if os.path.exists(output_path):
                print(f"跳过: '{pdf_filename}' 已存在。")
                skipped_count += 1
                continue

            try:
                print(f"正在转换: '{filename}' ...")
                # 转换单个文件
                convert(input_path, output_path)
                converted_count += 1
            except Exception as e:
                print(f"转换 '{filename}' 时出错: {e}")

        print(f"\n转换任务完成！")
        print(f"成功转换: {converted_count} 个文件")
        print(f"跳过已存在: {skipped_count} 个文件")

if __name__ == "__main__":
    batch_convert()
    # 转换完成后执行合并
    merge_pdfs()
