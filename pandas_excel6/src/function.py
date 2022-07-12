import pandas as pd
from log.log import *
from src import check_data
import numpy as np
import matplotlib.pyplot as plt
from src import to_excel


def operate_exam(df_sheet, index_axis, student_id):
    if index_axis == 'nan':
        logger.error('未找到班级列')
    else:
        # 取参数部分
        df_param = df_sheet.loc[:index_axis, '参数设置':].drop(index=index_axis)
        df_param.set_index('参数设置', inplace=True)
        df_param = df_param.dropna(axis=1, how='all')
        df_param.columns = df_param.loc['参数', :]
        df_param = df_param.drop(index='参数')
        df_param.set_index('题号', inplace=True)
        df_param = df_param.loc['小题分数':, :]
        df_param = df_param.fillna(value=0)           # 空值填0
        df_param.columns = range(1, len(df_param.columns)+1)

        # 取分数部分
        df_score = df_sheet.loc[index_axis:, :]
        df_score.columns = df_score.loc[index_axis]
        df_score = df_score.drop(index=index_axis)
        df_score.index = df_score.loc[:, '学生学号']
        df_score = df_score.drop(columns=['学生学号', '班级'])
        df_score = df_score.loc[student_id]
        df_score = df_score.dropna(axis=1, how='all')
        df_score.columns = range(1, len(df_score.columns)+1)
        return df_param, df_score

@logger.catch
def get_class_id(df_test_weight_index, sheet):
    k = -1
    class_id, everyclass_student, student_id = list(), list(), list()
    for i in df_test_weight_index:
        sheet_name = i
        df_sheet = pd.read_excel(sheet, sheet_name=str(i), header=1)
        # 查找表中含有的"班级"字符所在的坐标
        index_axis, columns_axis = check_data.find_it(df_sheet, '班级')
        k += 1
        if index_axis != 'nan' and columns_axis != 'nan':
            # 利用“班级”列分组
            class_sheet = df_sheet.loc[index_axis:]
            class_sheet.columns = class_sheet.loc[index_axis, :]
            class_sheet = class_sheet.drop(index=index_axis, axis=0)
            for name, group in class_sheet.groupby("班级"):
                class_id.append(name)
                group = group.loc[:, '学生学号'].values
                everyclass_student.append(list(group))
                student_id += list(group)
            break
    try:
        df_param, df_score = operate_exam(df_sheet, index_axis, student_id)
        return class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score
    except Exception as error:
        logger.debug('班级-学号处理异常')
        logger.exception(error)

@logger.catch
# 获取考核方式表格（随堂测验，作业，实验，考试）
def get_object_sheet(df_test_weight_index, sheet_name, sheet):
    # 存放客观评价dataframe
    object_evaluation_sheet = list()
    for i in df_test_weight_index:
            # print('页帧名称：', i)
            # 考核方式表
            df_test_sheet = pd.read_excel(sheet, sheet_name=str(i), index_col=0, header=1)
            object_evaluation_sheet.append(df_test_sheet)
    # print('页帧名称：', sheet_name)
    return object_evaluation_sheet

@logger.catch
# 获取主观评价表格（达成度 （教师评价法）, 达成度 （学生自评法））
def get_unobject_sheet(df_unobject_name, sheet):
    # 存放非客观评价dataframe
    unobject_evaluation_sheet = list()
    for i in df_unobject_name:
        # print('页帧名称：', i)
        df_unobject_sheet = pd.read_excel(sheet, sheet_name=str(i), header=None, index_col=0)
        unobject_evaluation_sheet.append(df_unobject_sheet)
    return unobject_evaluation_sheet

# 处理客观评价表
@logger.catch()
def operate_object_sheet(object_evaluation, df_test_weight, student_id, class_sheet, df_param, df_score):
    object_all_param, object_all_scores = list(), list()
    k = -1
    for i in range(len(object_evaluation)):
        if i != class_sheet:
            object_data = object_evaluation[i]
            # 学生成绩部分
            object_scores = object_data.loc[student_id]
            object_scores = object_scores.dropna(axis=1, how='all')
            try:
                object_data.columns = object_data.loc['参数'].drop(columns='参数')
                object_data.set_index('题号', inplace=True)
                # 参数部分
                object_data.columns = range(1, len(object_data.columns) + 1)
                object_parameter = object_data.filter(axis=0, regex=str('目标'))
                object_parameter_scores = object_data.loc['小题分数', :]
                object_parameter.columns = object_parameter_scores.columns = range(1, len(object_parameter.columns) + 1)
                object_parameter = object_parameter.T
                object_parameter.insert(0, '小题分数', object_parameter_scores.T)
                object_parameter = object_parameter.T.fillna(value=0)                     # 空值填0
                # 检查目标个数是否一致
                k += 1
                check_data.check_target_num(object_parameter, df_test_weight, df_test_weight.index.values[k])
                # 检查小题分数满分是否符合要求
                check_data.check_out(object_parameter, df_test_weight.index.values[k])
                object_all_param.append(object_parameter)
            except Exception as error:
                logger.exception(error)
            object_scores.columns = range(1, len(object_scores.columns) + 1)
            object_all_scores.append(object_scores)
        else:
            k += 1
            object_all_param.append(df_param)
            object_all_scores.append(df_score)
            # 检查小题分数满分是否符合要求
            check_data.check_out(df_param, df_test_weight.index.values[k])
    return object_all_param, object_all_scores

@logger.catch
# 处理非客观评价表
def operate_unobject_sheet(unobject_sheet, unobject_columns, student_id, df_index_values):
    unobject_param = list()
    unobject_scores = list()
    k = 0
    for i in unobject_sheet:
        # 参数部分
        single_param = i.loc['参数设置':'学生学号', :].drop(index=['学生学号', '参数设置'])
        single_param = single_param.dropna(axis=0, how='all')
        single_param = single_param.dropna(axis=1, how='any')
        single_param.columns = single_param.loc['参数', :]
        single_param.drop(index='参数', inplace=True)
        single_param.index = ['满分']
        unobject_param.append(single_param)
        check_data.check_object_out(single_param, df_index_values[k])
        k += 1

        # 分数部分
        single_unobject_scores = i.loc[student_id]
        single_unobject_scores.dropna(axis=1, how='all', inplace=True)
        single_unobject_scores.columns = unobject_columns
        unobject_scores.append(single_unobject_scores)
    return unobject_param, unobject_scores

def read_object_sheet(df_test_weight, sheet_name, student_id, k, df_param, df_score, sheet):
    # 获取客观评价表格\ 非客观评价部分表格
    object_evaluation = get_object_sheet(df_test_weight.index, sheet_name, sheet)
    # 对考核方式表进行处理,返回包含考核方式参数和分数的列表
    object_all_param, object_all_scores = operate_object_sheet(object_evaluation, df_test_weight, student_id, k, df_param, df_score)
    return object_all_param, object_all_scores

def read_unobject_sheet(df_index_values, sheet, student_id, df_test_weight):
    # 读取主观评价表
    unobject_evaluation = get_unobject_sheet(df_index_values, sheet)
    # 获取参数和分数部分
    unobject_param, unobject_scores = operate_unobject_sheet(unobject_evaluation, df_test_weight.columns, student_id, df_index_values)
    return unobject_param, unobject_scores

# 计算客观评价部分各目标总分和达成度
def count_object_evaluation(single_param, single_scores, df_test_weight):
    count_param = single_param.loc[df_test_weight.columns]
    # 计算各目标得分
    df_scores = single_scores.dot(count_param.T)
    for i in df_test_weight.columns.values:
        count_param.loc[str(i), '总分'] = (single_param.loc[str(i)] * single_param.loc['小题分数']).sum()
    count_param = count_param.loc[~(count_param == 0).all(axis=1)]
    # 达成度
    df_achieve = df_scores.div(count_param['总分'].T, fill_value=None)
    df_achieve = df_achieve.applymap(lambda x: '%.3f' % x)
    return df_achieve, df_scores

def count_object_target(object_all_scores, object_all_params, df_test_weight, sheet):
    # 保存得分、达成度、计算完总分的参数表参数表
    object_evaluation_out, object_evaluation_achieve = list(), list()
    for i in range(len(object_all_scores)):
        single_achieve, single_out = count_object_evaluation(object_all_params[i], object_all_scores[i], df_test_weight)
        object_evaluation_achieve.append(single_achieve)
        object_evaluation_out.append(single_out)

    return object_evaluation_achieve, object_evaluation_out

def concat_data(df_weight, object_evaluation_achieve, unobject_scores, student_id):
    all_achieve = list()
    for i in df_weight.index:
        all_achieve_data = pd.DataFrame(data=None, index=student_id)
        for j in object_evaluation_achieve:
            target_object_data = j.loc[:, str(i)]
            all_achieve_data = pd.merge(all_achieve_data, target_object_data, left_index=True, right_index=True)
        for k in unobject_scores:
            target_unobjectdata = k.loc[:, str(i)]
            all_achieve_data = pd.merge(all_achieve_data, target_unobjectdata, left_index=True, right_index=True)
        all_achieve_data.columns = df_weight.columns
        all_achieve.append(all_achieve_data)
    return all_achieve

def count_final_achieve(all_target_data, df_weight, student_id):
    # 计算达成度
    target_achieve = pd.DataFrame(index=student_id)
    for i in range(len(all_target_data)):
        count_weight = df_weight.loc[str(df_weight.index.values[i])]
        count_weight = count_weight[count_weight.values != 0]
        target_data = all_target_data[i].loc[:, count_weight.index]
        target_data = target_data.astype(float)
        achievement = round((target_data.mul(count_weight, axis=1)).sum(axis=1), 3)
        target_achieve.insert(len(target_achieve.columns), '课程目标' + str(i + 1), achievement)
    # 计算均值
    target_achieve.loc['平均值'] = target_achieve.apply(lambda x: round(x.mean(), 3), axis=0)
    return target_achieve

# 计算客观评价班级平均值和达成度
def count_class_object(out_data, object_param, class_id, everyclass_student, class_achieve, class_average, df_weight):
    score_part = object_param.loc[df_weight.index, :]
    object_param = (score_part * object_param.loc['小题分数']).T.sum()
    for i in everyclass_student:
        class_out = out_data.loc[i]
        average = round(class_out.mean(axis=0), 3)
        achieve = round(average.div(object_param), 3)
        class_average = pd.concat([class_average, average], axis=1, ignore_index=False)
        class_achieve = pd.concat([class_achieve, achieve], axis=1, ignore_index=False)
    class_achieve.dropna(axis=1, how='all', inplace=True)
    class_average.dropna(axis=1, how='all', inplace=True)
    class_achieve.columns = class_id
    class_average.columns = class_id
    return class_average.T, class_achieve.T

@logger.catch
# 计算非客观部分班级平均分和达成度
def count_class_unobject(unobject_scores, unobject_param, everyclass_student, class_achieve, class_average):
    try:
        for i in everyclass_student:
            scores = unobject_scores.loc[i]
            average = round(scores.mean(axis=0), 2)
            unobject_param.columns = unobject_scores.columns
            achieve = round(average.div(unobject_param), 3)
            class_average = pd.concat([class_average, average.T], axis=1, ignore_index=False)
            class_achieve = pd.concat([class_achieve, achieve.T], axis=1, ignore_index=False)
        return class_achieve, class_average
    except Exception as error:
        logger.debug('主观计算出现异常')
        logger.exception(error)

# 计算班级均值、达成度
def count_class_avg_ach(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student):
    class_ach = list()                      # 存放班级达成度
    class_avg = list()                      # 存放班级平均值
    # 客观评价部分
    class_objachieve = pd.Series(data=None, index=df_weight.index, dtype=object)
    class_objaverage = pd.Series(data=None, index=df_weight.index, dtype=object)
    for i in range(len(object_evaluation_out)):
        average, achieve = count_class_object(object_evaluation_out[i], object_all_param[i], class_id, everyclass_student, class_objachieve, class_objaverage, df_weight)
        class_ach.append(achieve)
        class_avg.append(average)
    class_unobjachieve = pd.Series(data=None, index=df_weight.index, dtype=object)
    class_unobjaverage = pd.Series(data=None, index=df_weight.index, dtype=object)
    # 主观部分
    for i in range(len(unobject_param)):
        class_achieve, class_average = count_class_unobject(unobject_scores[i], unobject_param[i], everyclass_student, class_unobjachieve, class_unobjaverage)
        class_achieve = class_achieve.dropna(axis=1, how='all')
        class_achieve.columns = class_id
        class_ach.append(class_achieve.T)
        class_average = class_average.dropna(axis=1, how='all')
        class_average.columns = class_id
        class_avg.append(class_average.T)
    return class_ach, class_avg

# 拼接各班达成度
def concat_class_achieve(class_ach, class_achieve, target_name, df_weight):
    for j in class_ach:
        class_achieve = pd.merge(class_achieve, j.loc[target_name], left_index=True, right_index=True)
    class_achieve.columns = df_weight.columns
    return class_achieve

all_everyclass_achieve = list()
def concat_everyclass_achieve(class_ach, class_id, df_weight):
    for i in class_id:
        class_achieve = pd.DataFrame(data=None, index=df_weight.index)
        class_achieve = concat_class_achieve(class_ach, class_achieve, i, df_weight)
        all_everyclass_achieve.append(class_achieve)
    return all_everyclass_achieve

def get_class_info(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student):
    class_ach, class_avg = count_class_avg_ach(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student)
    all_everyclass_achieve = concat_everyclass_achieve(class_ach, class_id, df_weight)
    return all_everyclass_achieve, class_avg, class_ach

# 统计部分
def static_data(achieve, student_id, df_weight):
    achieve_score = achieve.loc[student_id]
    average_data = achieve.loc['平均值']
    up_avg, up_single, dawn_single = list(), list(), list()
    up_avg_odd, up_single_odd, dawn_single_odd = list(), list(), list()
    static_data_single = {}
    static_data_single_odd = {}
    for i in achieve_score.columns:
        # 单项目标超过平均值人数
        count1 = (achieve_score.loc[student_id, str(i)] >= average_data.loc[str(i)]).sum()
        up_avg.append(count1)
        odd1 = ('%.2f%%') % ((count1 / len(student_id)) * 100)
        up_avg_odd.append(odd1)

        # 单项目标达成人数
        count2 = (achieve_score.loc[student_id, str(i)] >= 0.7).sum()
        up_single.append(count2)
        odd2 = ('%.2f%%') % ((count2 / len(student_id)) * 100)
        up_single_odd.append(odd2)

        # 单项目标未达成人数
        count3 = (achieve_score.loc[student_id, str(i)] < 0.7).sum()
        dawn_single.append(count3)
        odd3 = ('%.2f%%') % ((count3 / len(student_id)) * 100)
        dawn_single_odd.append(odd3)

    static_data_single['单项目标超过平均值人数'] = up_avg
    static_data_single['单项目标达成人数'] = up_single
    static_data_single['单项目标未达成人数'] = dawn_single
    static_data_single_odd['单项目标超过平均值人数'] = up_avg_odd
    static_data_single_odd['单项目标达成人数'] = up_single_odd
    static_data_single_odd['单项目标未达成人数'] = dawn_single_odd
    static_data_all = {}
    static_data_all_odd = {}
    # 所有目标均超过平均值
    count4 = 0
    for j in achieve_score.index.values:
        count = 0
        for i in achieve.columns:
            if achieve_score.loc[j, str(i)] >= average_data[str(i)]:
                count += 1
        if count == len(achieve.columns):
            count4 += 1
    odd4 = ('%.2f%%') % ((count4 / len(student_id)) * 100)
    static_data_all['所有目标超过平均值人数'] = count4
    static_data_all_odd['所有目标超过平均值人数'] = odd4
    # 所有目标均达成
    count5 = len(achieve_score.loc[(achieve_score >= 0.7).all(axis=1)])
    odd5 = ('%.2f%%') % ((count5 / len(student_id)) * 100)
    static_data_all['所有目标均达成人数'] = count5
    static_data_all_odd['所有目标均达成人数'] = odd5
    # 所有目标均未达成
    count6 = len(achieve_score.loc[(achieve_score < 0.7).all(axis=1)])
    odd6 = ('%.2f%%') % ((count6 / len(student_id)) * 100)
    static_data_all['所有目标均未达成人数'] = count6
    static_data_all_odd['所有目标均未达成人数'] = odd6

    static_data_all = pd.Series(static_data_all)
    static_data_single = pd.DataFrame(static_data_single)
    static_final_data = pd.concat([static_data_single.T, static_data_all.T])

    static_data_all_odd = pd.Series(static_data_all_odd)
    static_data_single_odd = pd.DataFrame(static_data_single_odd)

    static_data_single_odd = static_data_single_odd.T
    static_data_single_odd.columns = df_weight.index
    static_final_data.columns = df_weight.index
    return achieve, static_final_data, static_data_single_odd, static_data_all_odd

# 散点图
def diagram(target_achieve, student_id, df_weight, sheet, start_col):
    start_idx = 6       # 散点图开始写入的行
    for i in df_weight.index:
        achieve = target_achieve.loc[student_id, str(i)]
        l = len(student_id)
        fig = plt.figure()
        ax1 = fig.add_subplot(1, 1, 1)
        ax1.scatter(range(0, l), achieve, s=20, label='达成值')
        x1 = np.random.uniform(0, l, l)
        y1 = np.array([0.7] * l)
        x2 = np.random.uniform(0, l, l)
        y2 = np.array([target_achieve.loc['平均值', str(i)]] * l)
        ax1.plot(x1, y1, c='r', ls='--', label='达成值')
        ax1.plot(x2, y2, c='g', ls='--', label='平均值')
        plt.xlim(xmax=l, xmin=0)
        plt.ylim(ymax=1, ymin=0)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.title(str(i) + "达成度散点图")
        plt.legend()
        # plt.show()
        # 将散点图写入’达成度散点图‘页帧
        to_excel.plt_to_excel(fig, sheet, str(i) + "达成度散点图", start_col, start_idx)
        start_idx += 30