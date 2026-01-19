import radar
import os
import pandas as pd
import shutil


def format_column_list(columns):
    formatted_list = "\n".join([f"{idx}. {col}" for idx, col in enumerate(columns, 1)])
    return formatted_list


def get_user_input_column(df, column_purpose):
    while True:
        formatted_columns = format_column_list(df.columns[:4])
        print(f"\n可用的列名如下:\n{formatted_columns}")
        print("输入 'e' 退出程序。")
        user_input = input(f"\n请选择用于{column_purpose}的列名的编号: ").strip().lower()

        if user_input == 'e':
            return None

        try:
            column_number = int(user_input)
            if 1 <= column_number < len(df.columns[:4]):
                return df.columns[column_number - 1]
            else:
                print("输入的编号不正确，请重新输入！")
        except ValueError:
            print("输入无效，请输入正确的编号！")


def create_class_folders(grade, folder_path, df):
    class_column = get_user_input_column(df, "创建文件夹")

    if class_column is None:
        return  # 用户选择退出

    classes = df[class_column].unique()
    print(f"Y{grade}年级共有{len(classes)}个班级。")

    for class_name in classes:
        folder_name = os.path.join(folder_path, f'Y{grade}报告单', class_name)
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            print(f'已成功创建{class_name}文件夹')
        else:
            print(f"已存在{class_name}文件夹。")
    return class_column


def move_files_to_folders(df, folder_path, grade, class_column):
    for index, row in df.iterrows():
        class_name = row[class_column]
        doc_name = row['new_file_name'] + '.pdf'
        # doc_name = row['new_file_name'] + '.docx'
        src_path = os.path.join(folder_path, f'Y{grade}报告单', doc_name)
        dest_folder = os.path.join(folder_path, f'Y{grade}报告单', class_name)
        if not os.path.exists(dest_folder):
            print(f'找不到{class_name}文件夹。')
            continue
        # 检查文件是否存在并移动
        if os.path.exists(src_path):
            shutil.move(src_path, os.path.join(dest_folder, doc_name))
            print(f"{doc_name}已移动至{class_name}文件夹内。")
        else:
            print(f"找不到文件: {doc_name} ------ 请检查路径：{src_path}")


def main():
    result = radar.get_grade_and_folder_path()

    if result is None:
        print("用户选择退出程序。")
        return  # 直接退出 main 函数

    grade, folder_path = result
    xlsx_file = os.path.join(folder_path, f'Y{grade}_report.xlsx')

    if not os.path.exists(xlsx_file):
        print(f"找不到文件: {xlsx_file}")
        return

    df = pd.read_excel(xlsx_file)
    class_column = create_class_folders(grade, folder_path, df)

    if class_column is not None:
        move_files_to_folders(df, folder_path, grade, class_column)


if __name__ == '__main__':
    main()
