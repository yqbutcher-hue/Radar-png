import os
import pandas as pd
import radar


def rename_files_with_excel():
    grade, folder_path = radar.get_grade_and_folder_path()
    path = os.path.join(folder_path, f'Y{grade}_report.xlsx')
    if not os.path.exists(path):
        print(f'找不到该文件路径: {path}')
    df = pd.read_excel(path)

    # 检查new_file_name列是否存在
    if 'new_file_name' not in df.columns:
        print("Excel文件中缺少'new_file_name'列。")
        return

    for index, row in df.iterrows():
        old_file_path = os.path.join(folder_path, f'Y{grade}报告单', f'Y{grade}_{int(index) + 1}.docx')
        new_file_name = row['new_file_name']

        # 检查新文件名是否为空
        if pd.isnull(new_file_name):
            print(f"第{int(index) + 1}个学生的新文件名为空。")
            continue

        new_file_path = os.path.join(folder_path, f'Y{grade}报告单', f'{new_file_name}.docx')

        # 检查文件是否存在
        if not os.path.exists(old_file_path):
            print(f'找不到文件: {old_file_path}')
            continue
        if os.path.exists(new_file_path):
            print(f'文件已存在: {new_file_path}')
            continue

        try:
            os.rename(old_file_path, new_file_path)
            print(f'Renamed: {old_file_path} -> {new_file_path}')
        except Exception as e:
            print(f'无法重命名文件 {old_file_path} -> {new_file_path}: {e}')


def main():
    rename_files_with_excel()


if __name__ == '__main__':
    main()
