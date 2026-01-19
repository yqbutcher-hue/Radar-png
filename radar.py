import pandas as pd
import plotly.graph_objs as go
import plotly.io as pio
import os
import time as t


def get_grade_and_folder_path():
    while True:
        grade = input('请输入要操作的年级数字（如1、2、3）或者输入 e 退出程序：')

        if grade == 'e':
            return None

        if grade.isdigit() and 1 <= int(grade) <= 6:
            folder_path = os.path.join(os.path.dirname(__file__), f'Y{grade}_RADAR')
            if not os.path.exists(folder_path):
                print(f'找不到文件目录: {folder_path}')
                continue
            break
        else:
            print("输入的数字不符合要求，请输入小于或等于6的数字。")
    return grade, folder_path


def get_radar_dict():
    """
    返回包含雷达图配置的字典。
    """
    return {'subjects': ['chinese', 'math', 'english', 'physical', 'science'],
            'chinese_columns_1': ['汉语拼音', '识字写字', '阅读理解', '表达交流', '综合实践'],   # Y1
            'chinese_columns_2': ['识字写字', '阅读理解', '写话', '表达交流', '综合实践'],  # Y2
            'chinese_columns_3': ['识字写字', '阅读理解', '习作', '表达交流', '综合实践'],   # Y3
            'chinese_columns_4': ['识字写字', '阅读理解', '积累运用', '表达交流', '综合实践'],  # Y4 - Y6
            # 'chinese_columns_4': ['Literacy', 'Comprehension', 'Application', 'Communication', 'Activities'],
            'math_columns_1': ['计算能力', '审题能力', '解决问题', '实践操作', '双语能力'],
            'math_columns_2': ['审题能力', '计算能力', '知识应用', '问题解决', '双语能力'],
            # 'math_columns_2': ['Comprehension', 'Calculation', 'Application', 'Problem-Solving', 'Bilingualism'],
            'english_columns': ['Speaking', 'Reading', 'Writing', 'Use of English', 'Listening'],
            'physical_columns_1': ['BMI', '肺活量', '50米跑', '坐位体前屈', '1分钟跳绳'],
            'physical_columns_2': ['BMI', '肺活量', '50米跑', '坐位体前屈', '1分钟跳绳', '仰卧卷腹'],
            'physical_columns_3': ['BMI', '肺活量', '50米跑', '坐位体前屈', '1分钟跳绳', '仰卧卷腹', '50*8往返跑'],
            'science_columns_1': ['观察实验能力', '科学思维能力', '动手实践能力', '科学表达能力', '词汇应用能力'],
            }


def get_subject_choice(radar_dict):
    """
    显示所有学科供用户选择。
    """
    subjects = radar_dict['subjects']
    while True:
        print("\n选择学科：")
        for i, sub in enumerate(subjects, 1):
            print(f"{i}. {sub}")
        print("a. 全部学科")
        choice = input("输入学科序号或输入字母'a'选择全部学科：").strip().lower()

        if choice == 'a':
            return subjects  # 返回所有学科
        elif choice.isdigit() and 1 <= int(choice) <= len(subjects):
            return [subjects[int(choice) - 1]]  # 返回指定学科
        else:
            print("输入无效，请重新输入。")
        t.sleep(.5)


def get_student_choice(df):
    """
    提示用户输入学生姓名或选择生成全部学生的雷达图。
    """
    while True:
        print("\n选择学生：")
        print("a. 全部学生")
        choice = input("输入学生姓名或输入字母'a'生成全部学生的雷达图：")

        if choice == 'a':
            return df  # 返回全部学生的DataFrame
        elif choice in df['学生姓名'].values:
            return df[df['学生姓名'] == choice]  # 返回指定学生的DataFrame
        else:
            print(f"没有找到'{choice}'同学，请重新输入。")
        t.sleep(.5)


def select_columns_for_subject(sub, grade, radar_dict):
    """
    根据学科和年级选择适当地评价维度。
    """
    if sub == 'chinese':
        if grade in ['1']:
            return radar_dict['chinese_columns_1']
        elif grade in ['2']:
            return radar_dict['chinese_columns_2']
        elif grade in ['3']:
            return radar_dict['chinese_columns_3']
        else:
            return radar_dict['chinese_columns_4']
    elif sub == 'math':
        if grade in ['1', '2', '3']:
            return radar_dict['math_columns_1']
        else:
            return radar_dict['math_columns_2']
    elif sub == 'english':
        return radar_dict['english_columns']
    elif sub == 'physical':
        if grade in ['1', '2']:
            return radar_dict['physical_columns_1']
        elif grade in ['3', '4']:
            return radar_dict['physical_columns_2']
        else:
            return radar_dict['physical_columns_3']
    elif sub == 'science':
        if grade in ['1']:
            return radar_dict['science_columns_1']


def check_columns_exist(df, select_columns, sub):
    """
    检查DataFrame中是否存在指定的列。
    """
    missing_columns = [col for col in select_columns if col not in df.columns]
    if missing_columns:
        print(f"Excel中没有{sub}标题：\n{', '.join(missing_columns)}")
        return False
    return True


def create_radar_chart(sub, row, select_columns, output_folder):
    """
    为每个学生生成雷达图。
    """
    select_data = select_columns.copy()
    select_data.append(select_data[0])
    values = row[select_data].tolist()
    theta = select_data

    fig = go.Figure()
    # 添加雷达图的红色范围
    fig.add_trace(go.Scatterpolar(
        r=values,
        theta=theta,
        fill='toself',
        fillcolor='rgba(232, 208, 208, 0.5)',
        line=dict(width=4, color='#FF5722'),
        opacity=1,
        name='',
    ))

    # 添加圆形标记
    for i, val in enumerate(values[:-1]):  # 去除起始点
        fig.add_trace(go.Scatterpolar(
            r=[val],
            theta=[theta[i]],
            mode='markers',  # 设置为圆形节点
            marker=dict(
                size=14,  # 圆形节点的大小
                color='#FF5722',  # 与trace line的颜色一致
                opacity=1,
                line=dict(width=0, color='#FF5722'),  # 圆形节点的边界样式
            ),
            name='',
        ))

    vertical_offset = 8  # 自定义垂直偏移量，根据需要调整
    for i, val in enumerate(values):
        # 当学科为physical时，跳过不显示分数
        if sub != 'physical' or 'science':
            r = [val - vertical_offset]
            fig.add_trace(go.Scatterpolar(
                r=r,
                theta=[theta[i]],
                mode='text',
                text=str(round(val * 0.05, 1)),  # 保留整数部分
                textposition='bottom center',
                textfont=dict(family='PingFang SC', size=24, color='#FF5722'),
                name='',
            ))

    fig.update_polars(
        gridshape='circular',  #
        bgcolor='rgba(0,0,0,0)',  # 设置为透明背景
        radialaxis=dict(visible=True,
                        showgrid=True,
                        showticklabels=False,
                        gridwidth=2,
                        gridcolor='rgba(128, 128, 128, 0.2)',
                        linewidth=0,
                        range=[0, 100],
                        tickvals=[25, 50, 75, 100],  # 设置刻度值
                        nticks=4,  # 设置刻度线数量
                        ),
        angularaxis=dict(showgrid=True,  # 角度轴网格线
                         linecolor='rgba(128, 128, 128, 0.2)',
                         linewidth=2,
                         gridwidth=2,
                         gridcolor='rgba(128, 128, 128, 0.4)',
                         direction="clockwise",  # 设置方向为顺时针，使底边水平
                         tickfont=dict(family='PingFang SC', size=50, color='#393939'),
                         ),
    )

    fig.update_layout(
        showlegend=False,  # 不显示图例
        paper_bgcolor='rgba(0,0,0,0)',  # 整个图表的纸张背景设置为透明
        height=900,
        width=1300,
    )

    # 保存雷达图为.png文件
    student_name = row['学生姓名']
    output_file = os.path.join(output_folder, f"{student_name}.png")
    pio.write_image(fig, output_file, format="png")
    print(f"Saved {sub} {output_file}!")


def main():
    radar_dict = get_radar_dict()
    result = get_grade_and_folder_path()

    if result is None:
        print("用户选择退出程序。")
        return  # 直接退出 main 函数

    grade, folder_path = result
    xlsx_file = os.path.join(folder_path, f'Y{grade}_report.xlsx')
    # xlsx_file = os.path.join(folder_path, f'Y{grade}_report_english.xlsx')
    if not os.path.exists(xlsx_file):
        print(f"找不到文件: {xlsx_file}")
        return

    df = pd.read_excel(xlsx_file)
    selected_subjects = get_subject_choice(radar_dict)
    selected_students = get_student_choice(df)

    # 创建文件夹以保存雷达图
    for sub in selected_subjects:
        output_folder = os.path.join(folder_path, sub)
        os.makedirs(output_folder, exist_ok=True)
        select_columns = select_columns_for_subject(sub, grade, radar_dict)

        if not check_columns_exist(df, select_columns, sub):
            continue  # 如果列不存在，则跳过当前学科

        for _, row in selected_students.iterrows():
            create_radar_chart(sub, row, select_columns, output_folder)


if __name__ == '__main__':
    main()
