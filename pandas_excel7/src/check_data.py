import pandas as pd
from log.log import logger

# 定位表中things_needed字符所在的坐标
def find_it(dtx, things_needed):
    index_id = columns_id = 'nan'
    for j in range(len(dtx.index)):
        data = dtx.loc[j]
        if things_needed in data.values:
            index_id = j
            data.index = range(len(data))
            columns_id = data[data.values == str(things_needed)].index.values
    return index_id, columns_id[0]

# 检验封皮和课程目标达成情况报告
def check_cover(file):
    cover = pd.read_excel(file, sheet_name='封皮', index_col=0, header=None)
    report = pd.read_excel(file, sheet_name='课程目标达成情况报告', header=None)
    # 获取封皮数据
    cover.dropna(axis=0, how='all', inplace=True)
    cover.dropna(axis=1, how='all', inplace=True)
    cover.columns = range(len(cover.columns))
    cover.set_index(0, inplace=True)
    if '编写时间' in cover.index:
        cover.drop(index='编写时间', inplace=True)
    # 查找报告中的数据值
    for i in cover.index:
        # 遍历报告页帧获取封皮指定数据索引
        index_axis, columns_axis = find_it(report, str(i))
        if index_axis != 'nan' and columns_axis != 'nan':
            data = report.loc[index_axis, :]
            data = data.dropna()
            data.index = range(len(data))
            report_value = data[(data[data.values == str(i)].index[0]) + 1]
            # 比较封皮与报告页帧数据是否一致
            if report_value != cover.loc[str(i)].values:
                print('课程目标达成情况报告页帧【'+str(i)+'】数据与封皮页帧数据不一致，程序终止运行')
                logger.error('课程目标达成情况报告页帧【'+str(i)+'】数据与封皮页帧数据不一致，程序终止运行')
                exit()
        else:
            print('课程目标达成情况报告页帧不存在【'+str(i)+'】，程序终止运行')
            logger.error('课程目标达成情况报告页帧不存在【'+str(i)+'】，程序终止运行')
            exit()

# 检查权重和是否为1
def check_weight_sum(df_test_weight):
    weight_sum = df_test_weight.apply(lambda x: x.sum(), axis=0)
    if not (weight_sum.values == 1).all():
        print(weight_sum[weight_sum.values != 1].index[0]+'权重和数据不合法，程序终止运行')
        logger.debug('权重和数据不合法')
        logger.error(weight_sum[weight_sum.values != 1].index[0]+'权重和不为1，程序终止运行')
        exit()

# 目标个数是否一致、第二页帧权重设置是否和参数设置冲突、参数是否为1
def check_target_num(object_parameter, df_test_weight, sheet, weight_num, sheet_name):
    object_parameter = object_parameter.filter(axis=0, regex=str('目标'))
    obj_param = object_parameter[object_parameter == 1].count(axis=1)
    obj_num = object_parameter[object_parameter.sum(axis=1) != 0]
    weight_num = weight_num[weight_num != 0]
    if obj_param.sum() != 1 * obj_param.sum():
        logger.error(str(sheet_name) + '页帧参数设置异常，程序终止')
        print(str(sheet_name) + '页帧参数设置异常，程序终止')
        exit()
    if obj_num.index.tolist() != weight_num.index.tolist():
        logger.error(str(sheet_name) + '页帧权重设置和参数设置冲突，程序终止')
        print(str(sheet_name) + '页帧权重设置和参数设置冲突，程序终止')
        exit()
    if list(df_test_weight.columns.values) != list(object_parameter.index.values):
        print("课程目标个数设置不一致，程序终止")
        logger.debug("课程目标个数设置不一致，程序终止")
        if len(set(df_test_weight.columns.values)) > len(set(object_parameter.index.values)) - 1:
            logger.error(sheet + '页帧缺少索引:{value}', value=set(df_test_weight.columns.values) - set(object_parameter.index.values))
        else:
            logger.error(sheet + '页帧多出索引:{value}', value=set(object_parameter.index.values) - set(df_test_weight.columns.values))
        exit()

# 检查各目标得分是否超过单项总分
def check_out_error(parameter, scores, sheet_name):
    if '小题分数' in parameter.index:
        param_out = parameter.loc['小题分数', :]
    else:
        param_out = parameter.loc['满分', :]
    for i in scores.index:
        if scores.loc[i, :].values.tolist() > param_out.values.tolist():
            logger.error(sheet_name + '页帧学生分数设置异常，程序终止运行')
            print(sheet_name + '页帧学生分数设置异常，程序终止运行')
            exit()


# # 检查主观评价满分是否为1
# def check_object_out(object_out, sheet_name):
#     if object_out.values.all() != 1:
#         idx = object_out.loc[:, (object_out != 1).all(axis=0)]
#         print(sheet_name + "页帧参数" + str(idx.columns.tolist()) + "满分设置不合法, 程序终止")
#         logger.error(sheet_name + "页帧参数" + str(idx.columns.tolist()) + "满分设置不合法, 程序终止")
#         exit()

# 小题分数设置
# def check_out(df_param, sheet):
#     if df_param.loc['小题分数', :].sum() != 100:
#         print(sheet + "页帧小题分数设置不合法")
#         logger.error(sheet + "页帧小题分数设置不合法")
#         exit()