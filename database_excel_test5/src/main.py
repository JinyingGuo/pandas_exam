from src import get_info
from src import function
from log.log import *
from src import check_data
import os

sheet_name = get_info.get_excel_name()
for file in sheet_name:
    try:
        file_name = os.path.split(file)
        file_name = os.path.splitext(file_name[1])
        get_info.log_add(file_name[0])
        # check_data.check(file)
        # 提取课程目标达成情况报告信息
        df_weight, df_test_weight, df_index_values = get_info.get_weight(file)
        logger.info('\n 权重矩阵：{name}\n{}', df_weight, name='df_weight')
        # 处理班级，学号;拿到班级列页帧参数和分数部分
        class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score = get_info.operate_id(df_test_weight, file)
        # 计算达成度和均值
        target_achieve, object_evaluation_out, object_all_param, unobject_scores, unobject_param, object_all_scores, object_param, unobject_evaluation, all_target_data = get_info.read_sheet(df_test_weight, sheet_name, df_index_values, student_id, k, df_param, df_score, df_weight, file)
        # 计算班级达成度和平均值
        class_ach, class_avg = get_info.get_class_info(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student)
        # 数据统计部分
        achieve, static_final_data, static_data_single_odd, static_data_all_odd = function.static_data(target_achieve, student_id, df_weight)
        logger.info('\n 最终达成度：{name}\n{}', achieve, name='achieve')
        logger.info('\n 统计数据：{name}\n{}', static_final_data, name='static_final_data')
        logger.info('\n 统计数据：{name}\n{}', static_data_single_odd, name='static_data_single_odd')
        logger.info('\n 统计数据：{name}\n{}', static_data_all_odd, name='static_data_all_odd')
        # 散点图
        # function.diagram(target_achieve, student_id, df_weight)
        print('程序运行完成')
    except Exception as error:
        print(file_name[0], '表格出现异常')
