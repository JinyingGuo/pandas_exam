from src.main import *
from src import to_function
from src.function import *

# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

sheet = "../data/数据库原理及应用课程目标达成情况.xls"
name_sheet = df_weight.columns
data_name = r'E:\python\2021.9.17\database_excel_test\data\数据库原理及应用课程目标达成情况.xls'
sheet_test = pd.read_excel(data_name, sheet_name='随堂测验', header=2, index_col=1)
sheet_homework = pd.read_excel(data_name, sheet_name='作业', header=2, index_col=1)
sheet_experiment = pd.read_excel(data_name, sheet_name='实验', header=2, index_col=1)
sheet_exam = pd.read_excel(data_name, sheet_name='考试', header=3, index_col=2)
sheet_teacher = pd.read_excel(data_name, sheet_name='达成度 （教师评价法）', header=1, index_col=1)
sheet_student = pd.read_excel(data_name, sheet_name='达成度 （学生自评法）', header=1, index_col=1)

def get_ach_avg(option):

    df_test = sheet_test.loc[str(option), :]
    df_homework = sheet_homework.loc[str(option), :]
    df_experiment = sheet_experiment.loc[str(option), :]
    df_exam = sheet_exam.loc[str(option), :]
    df_teacher = sheet_teacher.loc[str(option)]
    df_teacher.index = df_teacher['参数设置']
    df_teacher = df_teacher.drop(columns='参数设置')
    df_student = sheet_student.loc[str(option)]
    df_student.index = df_student['参数设置']
    df_student = df_student.drop(columns='参数设置')

    lst_class_score = [df_test, df_homework, df_experiment, df_exam, df_teacher, df_student]
    return lst_class_score

def get_everysheet_scores():
    sheet_test = pd.read_excel(data_name, sheet_name='随堂测验', header=1, index_col=0)
    sheet_homework = pd.read_excel(data_name, sheet_name='作业', header=1, index_col=0)
    sheet_experiment = pd.read_excel(data_name, sheet_name='实验', header=1, index_col=0)
    sheet_exam = pd.read_excel(data_name, sheet_name='考试', header=1, index_col=1)
    # sheet_teacher = pd.read_excel(data_name, sheet_name='达成度 （教师评价法）', header=2, index_col=0)
    # sheet_student = pd.read_excel(data_name, sheet_name='达成度 （学生自评法）', header=2, index_col=0)
    df_test = sheet_test.loc[student_id, '随堂测验得分':]
    df_homework = sheet_homework.loc[student_id, '作业得分':]
    df_experiment = sheet_experiment.loc[student_id, '实验得分':]
    df_exam = sheet_exam.loc[student_id, '支撑课程目标总分值':]
    # df_teacher = sheet_teacher.loc[student_id, '目标1':]
    # df_student = sheet_student.loc[student_id, '目标1':]

    lst_excel_scores = [df_test, df_homework, df_experiment, df_exam]
    return lst_excel_scores

# 各表格平均值
def get_class_avg():
    # 提取excel平均值
    lst_class_score = get_ach_avg('平均分')

    for i in range(len(lst_class_score) - 2):
        lst_class_score[i] = lst_class_score[i].dropna(axis=1, how='all')
        lst_class_score[i].index = class_id
        if i != 3:
            lst_class_score[i] = lst_class_score[i].drop(columns='参数')
    return lst_class_score, class_avg, name_sheet

# 各表格达成度
def get_class_ach():
    # 提取excel达成度
    lst_class_score = get_ach_avg('达成度')
    for i in range(len(lst_class_score) - 2):
        lst_class_score[i] = lst_class_score[i].dropna(axis=1, how='all')
        lst_class_score[i].index = class_id
        if i != 3:
            lst_class_score[i] = lst_class_score[i].drop(columns='参数')
    return lst_class_score, class_ach, name_sheet

# 各目标得分
def get_scores():
    # 计算数据
    object_evaluation_achieve, object_evaluation_out, object_all_param = to_function.count_object_target(object_all_scores, object_param, df_weight)
    lst_count_scores = list()
    for i in range(len(object_evaluation_out)):
        score_achieve = pd.merge(object_evaluation_out[i], object_evaluation_achieve[i], left_index=True, right_index=True)
        lst_count_scores.append(score_achieve)

    lst_scores = lst_count_scores

    # excel数据
    lst_excel_scores = get_everysheet_scores()
    return lst_scores, lst_excel_scores, name_sheet

# 达成度和均值对比
def get_final_achieve():
    # 计算出的达成度和均值
    count_achieve = to_function.count_final_achieve(all_target_data, df_weight, student_id)

    # excel数据
    excel_achieve = pd.read_excel(data_name, sheet_name='达成度散点图', header=2, index_col=0)
    excel_achieve = excel_achieve.filter(regex='达成度')
    excel_achieve.columns = df_weight.index
    excel_achieve = excel_achieve.loc[1:'平均值', :]
    excel_achieve.index = count_achieve.index
    return count_achieve, excel_achieve

# 获取统计部分的数据

def static_data():
    static_excel_data = pd.read_excel(data_name, sheet_name='达成度散点图', header=2, index_col=0)
    # excel数据
    static_excel_data = static_excel_data.loc['单项目标超过平均值人数':, :]
    static_excel_data = static_excel_data.dropna(axis=1, how='all')
    static_excel_data.columns = range(len(static_excel_data.columns))
    static_excel_data.loc[:, 1] = static_excel_data.loc[:, 1].apply(lambda x: format(x, '.2%'))
    static_excel_data.loc[:'单项目标未达成人数', 3] = static_excel_data.loc[:'单项目标未达成人数', 3].apply(lambda x: format(x, '.2%'))
    static_excel_data.loc[:'单项目标未达成人数', 5] = static_excel_data.loc[:'单项目标未达成人数', 5].apply(lambda x: format(x, '.2%'))
    static_excel_data.loc[:'单项目标未达成人数', 7] = static_excel_data.loc[:'单项目标未达成人数', 7].apply(lambda x: format(x, '.2%'))
    return static_final_data, static_excel_data

def all_data_contrast():
    # 将计算数据拼接为达成度散点图
    contrast_data = pd.DataFrame(data=None, index=student_id)
    for i in range(len(all_target_data)):
        single_target = pd.merge(all_target_data[i], target_achieve[df_weight.index[i]], left_index=True, right_index=True)
        single_target = single_target.astype(float)
        single_target = single_target.dropna(axis=1, how='all')
        single_target.insert(len(single_target.columns), '平均值', target_achieve.loc['平均值', df_weight.index[i]])
        contrast_data = pd.merge(contrast_data, single_target, left_index=True, right_index=True)
    count_data = contrast_data.applymap("{0:.03f}".format)
    count_data.columns = range(len(count_data.columns))

    # 取excel数据
    excel_data = pd.read_excel(sheet, sheet_name='达成度散点图', index_col=0)
    excel_data = excel_data.loc[1:138].applymap("{0:.03f}".format)
    excel_data.drop(excel_data.columns[len(excel_data.columns) - 1], axis=1, inplace=True)
    excel_data.columns = range(len(excel_data.columns))
    excel_data.index = student_id
    return count_data, excel_data