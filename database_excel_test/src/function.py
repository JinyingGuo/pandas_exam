import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from log.log import *

# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

# 定位表中含有的"班级"字符所在的坐标
def find_it(dtx, things_needed):
    index_id = 'nan'
    columns_id = 'nan'
    for j in range(len(dtx.index)):
        data = dtx.loc[j]
        if things_needed in data.values:
            index_id = j
            data.index = range(len(data))
            columns_id = data[data.values == str(things_needed)].index.values
    return index_id, columns_id[0]

def operate_exam(df_sheet, index_axis, student_id):
    # 取参数部分
    df_param = df_sheet.loc[:index_axis, '参数设置':'总分'].drop(index=index_axis, columns='总分')
    df_param.index = df_param.loc[:, '参数设置']
    df_param = df_param.drop(columns='参数设置')
    df_param = df_param.dropna(axis=1, how='all')
    df_param.columns = df_param.loc['参数', :]
    df_param = df_param.drop(index='参数')
    df_param.index = df_param['题号']
    df_param = df_param.drop(columns=['题号'])
    df_param = df_param.loc['小题分数':, :]
    df_param = df_param.fillna(value=0)  # 空值填0
    df_param.columns = range(1, len(df_param.columns)+1)

    # 取分数部分
    df_score = df_sheet.loc[index_axis:, :'总分'].drop(columns='总分')
    df_score.columns = df_score.loc[index_axis]
    df_score = df_score.drop(index=index_axis)
    df_score.index = df_score.loc[:, '学生学号']
    df_score = df_score.drop(columns=['学生学号', '班级'])
    df_score = df_score.loc[student_id]
    df_score = df_score.dropna(axis=1, how='all')
    df_score.columns = range(1, len(df_score.columns)+1)
    return df_param, df_score

def get_class_id(df_test_weight_index, sheet):
    k = -1
    class_id, everyclass_student, student_id = list(), list(), list()
    for i in df_test_weight_index:
        sheet_name = i

        df_sheet = pd.read_excel(sheet, sheet_name=str(i), header=1)
        # 查找表中含有的"班级"字符所在的坐标
        index_axis, columns_axis = find_it(df_sheet, '班级')
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
    df_param, df_score = operate_exam(df_sheet, index_axis, student_id)
    return class_id, everyclass_student, student_id, sheet_name, k, df_param, df_score

# 获取考核方式表格（随堂测验，作业，实验，考试）
def get_object_sheet(df_test_weight_index, sheet_name, sheet):
    # 存放客观评价dataframe
    object_evaluation_sheet = list()
    for i in df_test_weight_index:
        if i != sheet_name:
            print('页帧名称：', i)
            # 考核方式表
            df_test_sheet = pd.read_excel(sheet, sheet_name=str(i), index_col=0, header=1)
            object_evaluation_sheet.append(df_test_sheet)
    print('页帧名称：', sheet_name)
    return object_evaluation_sheet

# 获取非客观评价部分表格（达成度 （教师评价法）, 达成度 （学生自评法））
def get_unobject_sheet(df_unobject_name, sheet_name, sheet):
    # 存放非客观评价dataframe
    unobject_evaluation_sheet = list()
    for i in df_unobject_name:
        print('页帧名称：', i)
        if i != sheet_name:
            df_unobject_sheet = pd.read_excel(sheet, sheet_name=str(i), header=None, index_col=0)
            unobject_evaluation_sheet.append(df_unobject_sheet)
    return unobject_evaluation_sheet

# 处理客观评价表
def operate_object_sheet(object_evaluation, df_test_weight, student_id):
    object_all_param = list()
    object_all_scores = list()

    for i in object_evaluation:
        object_data = i
        # 处理表格无用数据
        object_data = object_data.loc[:, :'总分'].drop(columns=['总分'])
        # 学生成绩部分
        object_scores = object_data.loc[student_id]
        object_scores = object_scores.dropna(axis=1, how='all')
        object_data.columns = object_data.loc['参数'].drop(columns='参数')
        object_data.index = object_data['题号']
        object_data = object_data.drop(columns=['题号'])
        # 参数部分
        try:
            object_parameter = object_data.loc['小题分数':list(df_test_weight.columns.values)[-1]]
            object_parameter = object_parameter.fillna(value=0)  # 空值填0
            object_parameter.columns = range(1, len(object_parameter.columns) + 1)
            object_scores.columns = range(1, len(object_scores.columns) + 1)
            object_all_param.append(object_parameter)
            object_all_scores.append(object_scores)
        except Exception as error:
           logger.exception('出现异常', error)
    return object_all_param, object_all_scores

# 将考试页帧放入list
def back_list(object_all_param, object_all_scores, df_test_weight, k, df_param, df_score):
    if k <= len(df_test_weight.index):
        object_all_param.insert(k, df_param)
        object_all_scores.insert(k, df_score)
    return object_all_param, object_all_scores

# 计算客观评价部分各目标总分和达成度
def count_object_evaluation(single_param, single_scores, df_weight, df_weight_index):
    count_param = single_param.loc[df_weight.index]
    # 计算各目标得分
    df_scores = single_scores.dot(count_param.T)
    for i in df_weight_index:
        count_param.loc[str(i), '总分'] = (single_param.loc[str(i)] * single_param.loc['小题分数']).sum()

    count_param = count_param.loc[~(count_param == 0).all(axis=1)]
    df_achieve = df_scores.div(count_param['总分'].T, fill_value=None)
    df_achieve = df_achieve.applymap(lambda x: '%.3f' % x)
    return df_achieve, df_scores, count_param

# 处理非客观评价表
def operate_unobject_sheet(unobject_sheet, unobject_columns, student_id):
    unobject_param = list()
    unobject_scores = list()
    for i in unobject_sheet:
        # 参数部分
        single_param = i.loc['参数设置':'学生学号', :].drop(index=['学生学号', '参数设置'])
        single_param = single_param.dropna(axis=0, how='all')
        single_param = single_param.dropna(axis=1, how='any')
        single_param.columns = single_param.loc['参数', :]
        single_param.drop(index='参数', inplace=True)
        single_param.index = ['满分']
        unobject_param.append(single_param)

        # 分数部分
        single_unobject_scores = i.loc[student_id]
        single_unobject_scores.dropna(axis=1, how='all', inplace=True)
        single_unobject_scores.columns = unobject_columns
        unobject_scores.append(single_unobject_scores)
    return unobject_param, unobject_scores

# 拼接数据
all_achieve = list()
def count_object_evaluate(object_evaluation_achieve, unobject_scores, target_name, all_target_data, df_weight):
    for i in object_evaluation_achieve:
        target_object_data = i.loc[:, str(target_name)]
        all_target_data = pd.merge(all_target_data, target_object_data, left_index=True, right_index=True)
    # print(all_target_data)
    for i in unobject_scores:
        target_unobjectdata = i.loc[:, str(target_name)]
        all_target_data = pd.merge(all_target_data, target_unobjectdata, left_index=True, right_index=True)
    all_target_data.columns = df_weight.columns
    all_achieve.append(all_target_data)
    return all_achieve

# 计算达成度和均值
def count_evaluate(target_data, df_weight, target_name):
    df_weight = df_weight.loc[str(target_name)]
    df_weight = df_weight[df_weight.values != 0]
    target_data = target_data.loc[:, df_weight.index.values]
    target_data = target_data.astype(float)
    achievement = round((target_data.mul(df_weight, axis=1)).sum(axis=1), 3)
    return achievement


# 计算客观评价班级平均值和达成度
def count_class_object(out_data, object_param, class_id, everyclass_student, class_achieve, class_average):
    out_data = out_data.loc[:, object_param.index]
    object_param = object_param.loc[:, '总分']
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


# 计算非客观部分班级平均分和达成度
def count_class_unobject(unobject_scores, unobject_param, everyclass_student, class_achieve, class_average):
    for i in everyclass_student:
        scores = unobject_scores.loc[i]
        average = round(scores.mean(axis=0), 2)
        unobject_param.columns = unobject_scores.columns
        achieve = round(average.div(unobject_param), 3)
        class_average = pd.concat([class_average, average.T], axis=1, ignore_index=False)
        class_achieve = pd.concat([class_achieve, achieve.T], axis=1, ignore_index=False)
    return class_achieve, class_average

# 拼接各班达成度
def concat_class_achieve(class_ach, class_achieve, target_name, df_weight):
    for j in class_ach:
        class_achieve = pd.merge(class_achieve, j.loc[target_name], left_index=True, right_index=True)
    class_achieve.columns = df_weight.columns
    return class_achieve

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
def diagram(target_achieve, student_id, df_weight):
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
        plt.show()

