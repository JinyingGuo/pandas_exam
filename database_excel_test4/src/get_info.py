import pandas as pd
from src import function
from src import to_function
from log.log import *
import os

def get_excel_name():
    filelist = []
    for root, dirs, files in os.walk('..\data'):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                filelist.append(file)
        file_key = list(range(1, len(filelist) + 1))
        path = dict(zip(file_key, filelist))
        path[len(path) + 1] = "处理全部excel"
        for key, value in path.items():
            print(key, value)
        print("请输入想要处理的excel表格序号:")
        a = int(input())
        filelist.clear()
        if a == len(path):
            for file in files:
                file_name = os.path.join(root, file)
                filelist.append(file_name)
        else:
            filelist.append(os.path.join(root, path[a]))
    return filelist


def get_data(sheet):
    df = pd.read_excel(sheet, sheet_name='课程目标达成情况报告', index_col=0)
    df1 = pd.read_excel(sheet, sheet_name=None)

    # 获取客观评价权重部分
    try:
        df_test_weight = df[~df.index.duplicated(keep='first')]      # 去除课程目标达成情况评价报告表的重复索引
        df_test_weight = df_test_weight.loc['考核方式':'权重和', :]

        df_test_weight = df_test_weight.dropna(axis=1, how='all')
        df_test_weight.columns = df_test_weight.loc['考核方式']
        df_test_weight = df_test_weight.drop(index=['考核方式', '权重和'])
    except Exception as error:
        logger.exception('出现异常', error)
    # print('客观评价权重：\n', df_test_weight)

    # 获取需要读取主观评价页帧的表名
    try:
        df_unobject_index = df[~df.index.duplicated(keep='last')]
        df_unobject_index.index = map(lambda x: str(x).replace('\n', ' '), [x for x in df_unobject_index.index])   # 处理掉索引中的 /n
        df_unobject_index = df_unobject_index.loc['权重和':, :]
        # 保留excel中存在的页帧
        df_index = [x for x in df_unobject_index.index if x in list(df1)]
        # 获取除考核方式外的页帧名
        df_index_values = [x for x in df_index if x not in list(df_test_weight.index)]
    except Exception as error:
        logger.exception('出现异常', error)
    # 拿到三个评价方式权重
    data_weight = df_unobject_index.loc['课程目标达成度情况':]
    data_weight.columns = data_weight.loc['课程目标达成度情况']
    data_weight = data_weight.drop(index='课程目标达成度情况')
    data_weight = data_weight.loc[:, '权重值':]
    data_weight = data_weight.dropna(axis=0, how='any')
    data_weight = data_weight.loc[:, '权重值']
    # print(data_weight)
    return df_test_weight, df_index_values, data_weight


# 处理权重值表
def operate_weight(df_test_weight, data_weight):

    df_weight = df_test_weight.T
    df_weight = df_weight * data_weight['达成度 （考核成绩法）']
    data_weight = data_weight.drop(index='达成度 （考核成绩法）')

    for i in data_weight.index:
        df_weight[i] = data_weight[i]
    # print('权重值表：\n', df_weight)
    return df_weight

def get_weight(sheet):
    # 提取课程目标达成情况报告信息
    df_test_weight, df_index_values, data_weight = get_data(sheet)
    # 处理权重矩阵
    df_weight = operate_weight(df_test_weight, data_weight)
    return df_weight, df_test_weight, df_index_values

def operate_id(df_test_weight, sheet):
    class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score = function.get_class_id(df_test_weight.index, sheet)
    return class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score

# 读取页帧
def read_sheet(df_test_weight, sheet_name, df_index_values, student_id, k, df_param, df_score, df_weight, sheet):
    try:
        # 获取客观评价表格\ 非客观评价部分表格
        object_evaluation = function.get_object_sheet(df_test_weight.index, sheet_name, sheet)
        unobject_evaluation = function.get_unobject_sheet(df_index_values, sheet_name, sheet)
        # 对考核方式表进行处理,返回包含考核方式参数和分数的列表
        object_all_param, object_all_scores = function.operate_object_sheet(object_evaluation, df_test_weight, student_id)
        # 将包含班级列页帧放入list
        object_param, object_all_scores = function.back_list(object_all_param, object_all_scores, df_test_weight, k,
                                                             df_param, df_score)
        # 计算客观评价各目标总分和达成度
        object_evaluation_achieve, object_evaluation_out, object_all_param = to_function.count_object_target(
            object_all_scores, object_param, df_weight)
        # 对两个非客观评价表进行处理
        unobject_param, unobject_scores = function.operate_unobject_sheet(unobject_evaluation, df_test_weight.columns,
                                                                          student_id)
        # 拼接计算达成度、均值所需数据
        all_target_data = to_function.concat_data(df_weight, object_evaluation_achieve, unobject_scores, student_id)
        # 计算达成度和均值
        target_achieve = to_function.count_final_achieve(all_target_data, df_weight, student_id)
        return target_achieve, object_evaluation_out, object_all_param, unobject_scores, unobject_param, object_all_scores, object_param, unobject_evaluation, all_target_data
    except Exception as error:
        logger.exception('出现异常', error)


def get_class_info(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student):
    class_ach, class_avg = to_function.count_class_avg_ach(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student)
    all_everyclass_achieve = to_function.concat_everyclass_achieve(class_ach, class_id, df_weight)
    return all_everyclass_achieve, class_avg
