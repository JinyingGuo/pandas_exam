import pandas as pd
from log.log import *
from src import check_data
from src import function
from src import to_excel
import os

def get_excel_name():
    filelist = []
    for root, dirs, files in os.walk('..\data'):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                filelist.append(file)
        path = dict(zip(list(range(1, len(filelist) + 1)), filelist))
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
        # 检查权重和是否为1
        check_data.check_weight_sum(df_test_weight)
    except Exception as error:
        # logger.info('出现异常')
        logger.exception(error)

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
        # logger.info('出现异常', error)
        logger.exception(error)
    # 拿到三个评价方式权重
    try:
        data_weight = df_unobject_index.loc['课程目标达成度情况':]
        data_weight.columns = data_weight.loc['课程目标达成度情况']
        data_weight = data_weight.drop(index='课程目标达成度情况')
        data_weight = data_weight.loc[:, '权重值':]
        data_weight = data_weight.dropna(axis=0, how='any')
        data_weight = data_weight.loc[:, '权重值']
        return df_test_weight, df_index_values, data_weight
    except Exception as error:
        logger.exception(error)

# 处理权重值表
def operate_weight(df_test_weight, data_weight):
    df_weight = df_test_weight.T
    df_weight = df_weight * data_weight['达成度 （考核成绩法）']
    data_weight = data_weight.drop(index='达成度 （考核成绩法）')
    for i in data_weight.index:
        df_weight[i] = data_weight[i]
    return df_weight

def get_weight(sheet):
    # 提取课程目标达成情况报告信息
    df_test_weight, df_index_values, data_weight = get_data(sheet)
    # 处理权重矩阵
    df_weight = operate_weight(df_test_weight, data_weight)
    return df_weight, df_test_weight, df_index_values, data_weight

def operate_sheet(df_test_weight, sheet, df_index_values, df_weight, data_weight):
    # 处理班级列
    class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score = function.get_class_id(df_test_weight.index, sheet)
    # 获取客观评价参数和分数部分
    object_all_params, object_all_scores = function.read_object_sheet(df_test_weight, sheet_name, student_id, k, df_param, df_score, sheet)
    # 获取主观评价参数和分数部分
    unobject_params, unobject_scores = function.read_unobject_sheet(df_index_values, sheet, student_id, df_test_weight)
    # 计算客观部分得分和达成度
    object_evaluation_achieve, object_evaluation_out = function.count_object_target(object_all_scores, object_all_params, df_test_weight, sheet)
    # 拼接计算达成度和均值所需数据
    all_target_data = function.concat_data(df_weight, object_evaluation_achieve, unobject_scores, student_id)
    # 计算达成度和均值
    final_target_achieve = function.count_final_achieve(all_target_data, df_weight, student_id)
    # 计算班级均值和达成度(三个班级的达成度和均值列表)
    all_everyclass_achieve, class_avg, class_ach = function.get_class_info(object_evaluation_out, object_all_params, unobject_scores, unobject_params, df_weight, class_id, everyclass_student)
    # 统计部分
    achieve, static_final_data, static_data_single_odd, static_data_all_odd = function.static_data(final_target_achieve, student_id, df_weight)
    logger.info('\n 最终达成度：{name}\n{}', achieve, name='achieve')
    logger.info('\n 统计数据：{name}\n{}', static_final_data, name='static_final_data')
    logger.info('\n 统计数据：{name}\n{}', static_data_single_odd, name='static_data_single_odd')
    logger.info('\n 统计数据：{name}\n{}', static_data_all_odd, name='static_data_all_odd')
    # 将数据写入表格
    start_col, new_path = to_excel.write_excel(object_all_params, object_all_scores, object_evaluation_achieve,
                                            object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k,
                                            student_id, everyclass_student, class_id, df_index_values, all_target_data,
                                            df_weight, data_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd)

    # 散点图
    function.diagram(final_target_achieve, student_id, df_weight, new_path, start_col)
    return final_target_achieve
