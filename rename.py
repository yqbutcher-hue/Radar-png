import os
import pandas as pd
import radar


def rename_files_with_excel():
    grade, folder_path = radar.get_grade_and_folder_path()
    path = os.path.join(folder_path, f'Y{grade}.xlsx')
    if not os.path.exists(path):
        print(f'找不到该文件路径: {path}')
    df = pd.read_excel(path)

    # 检查new_file_name列是否存在
    if '姓名' not in df.columns:
        print("Excel文件中缺少'姓名'列。")
        return

    for index, row in df.iterrows():
        old_file_path = os.path.join(folder_path, f'Y{grade}报告单', f'Y{grade}报告单_{int(index) + 1}.docx')
        # 新文件名为 班级_姓名
        class_name = str(row['班级'])
        new_file_name = f"{class_name}_{row['姓名']}"

        # 创建班级目录（如果不存在）
        class_dir = os.path.join(folder_path, f'Y{grade}报告单', class_name)
        if not os.path.exists(class_dir):
            os.makedirs(class_dir)

        # 新文件路径为 班级目录/班级_姓名.docx
        new_file_path = os.path.join(class_dir, f'{new_file_name}.docx')

        # 检查新文件名是否为空
        if pd.isnull(new_file_name):
            print(f"第{int(index) + 1}个学生的新文件名为空。")
            continue

        # ...新路径已在上方定义...

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
