import os
from docx2pdf import convert
import radar


def convert_docx_to_pdf(input_path, output_path):
    try:
        convert(input_path, output_path)
        print(f"{os.path.basename(input_path)} 转换成功.")
    except Exception as e:
        print(f"{os.path.basename(input_path)} 转换失败: {e}")


def get_files_to_convert(folder_path):
    return [f for f in os.listdir(folder_path) if f.endswith(".docx")]


def convert_to_pdf(grade, folder_path):
    folder_path = os.path.join(folder_path, f'Y{grade}报告单')
    if not os.path.exists(folder_path):
        print(f'找不到文件目录: {folder_path}')
        return

    files = get_files_to_convert(folder_path)
    if not files:
        print("没有找到Word文档！")
        return

    for file in files:
        input_path = os.path.join(folder_path, file)
        output_path = os.path.splitext(input_path)[0] + ".pdf"
        convert_docx_to_pdf(input_path, output_path)


def main():
    result = radar.get_grade_and_folder_path()

    if result is None:
        print("用户选择退出程序。")
        return  # 直接退出 main 函数

    grade, folder_path = result
    convert_to_pdf(grade, folder_path)


if __name__ == "__main__":
    main()
