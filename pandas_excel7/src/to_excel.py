from log.log import *
import pandas as pd
import xlwings as xw
import os

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# 定位表中things_needed字符所在的坐标
def find_it(dtx, things_needed):
    index_id = 'nan'
    for j in range(len(dtx.index)):
        data = dtx.loc[j]
        if things_needed in data.values:
            index_id = j
            data.index = range(len(data))
    return index_id

def write_excel(object_all_params, object_all_scores, object_evaluation_achieve,
                                            object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k,
                                            student_id, everyclass_student, class_id, df_index_values, all_target_data, df_weight, data_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(sheet)
    # app.display_alerts = False
    # 将评价页帧计算数据写入excel
    object_all_out = write_out_achieve(wb, object_all_params, object_all_scores, object_evaluation_achieve,
                                            object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k, everyclass_student, class_id, df_index_values)
    # # 添加‘达成度散点图’页帧
    start_col = final_data_to_excel(wb, all_target_data, student_id, df_index_values, df_weight, data_weight, df_test_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd)
    # # 处理‘课程目标达成情况报告’页帧
    write_reporter(wb, sheet, df_test_weight, class_ach, class_id, df_index_values, data_weight)
    file_name = os.path.split(sheet)
    file_name = os.path.splitext(file_name[1])
    new_path = '../data_new/' + str(file_name[0] + '.xls')
    wb.save(new_path)
    wb.close()
    app.quit()
    app.kill()
    return start_col, new_path, object_all_out

# 处理评价页帧数据写入
def write_out_achieve(wb, object_all_params, object_all_scores, object_evaluation_achieve, object_evaluation_out,
                      df_test_weight, class_ach, class_avg, sheet, k, everyclass_student, class_id, df_index_values):
    try:
        sheet_name = df_test_weight.index  # 客观评价页帧名称
        object_all_out = list()
        for i in range(len(sheet_name)):
            df_old = pd.read_excel(sheet, sheet_name=str(sheet_name[i]), header=1)
            sht = wb.sheets[str(sheet_name[i])]  # 写入数据页帧
            if '总分' in df_old.columns:
                df_old = df_old.loc[:, :'总分']
                col_nums = len(df_old.columns)
            else:
                df_old = df_old.loc[:, :'总分']
                col_nums = len(df_old.columns) + 1
            # 处理总分参数部分
            param_out = param_to_excel(sht, df_test_weight, object_all_params[i], col_nums, i, k)
            # 将客观评价班级平均分、达成度写入excel
            object_class_data(df_old, class_ach[i], class_avg[i], everyclass_student, col_nums + 1, class_id, sht)
            # 学生总分、得分、达成度
            end_col, object_all_out = student_data_to_excel(sht, df_old, object_all_scores[i], col_nums, object_evaluation_out[i], everyclass_student, sheet_name[i], param_out, object_evaluation_achieve[i], df_test_weight, object_all_out)
            # 将写入数据部分设置居中和框线
            excel_style(sht, col_nums, end_col)
        # 将主观评价班级平均分、达成度写入表格
        average = class_avg[-len(df_index_values):]
        achieve = class_ach[-len(df_index_values):]
        for i in range(len(df_index_values)):
            sht = wb.sheets[df_index_values[i]]
            df_sheet = pd.read_excel(sheet, sheet_name=df_index_values[i], header=None)
            unobject_class_data(sht, df_sheet, average[i], achieve[i], everyclass_student, class_id)
        return object_all_out
    except Exception as error:
        logger.exception(error)

# 客观评价页帧参数部分写入
def param_to_excel(sht, df_test_weight, object_param, col_nums, i, k):
    # 总分
    object_out = pd.DataFrame((object_param.loc[df_test_weight.columns, :] * object_param.loc['小题分数']).T.sum())
    out = object_param.loc['小题分数'].T.sum()
    object_out = object_out.T
    object_out.insert(0, '小题分数', out)
    sht[3, col_nums-1].options(transpose=True).value = object_out.values.tolist()
    return object_out.T

# 将班级平均分、达成度写入excel
def object_class_data(df_sheet, class_ach, class_avg, everyclass_student, col_nums, class_id, sht):
    for i in range(len(everyclass_student)):
        index_id = find_it(df_sheet, everyclass_student[i][-1]) + 2
        average = class_avg.loc[:, ~(class_avg == 0).all(axis=0)]
        achieve = class_ach.dropna(axis=1, how='all')
        sht[index_id + 1, col_nums-1].value = average.loc[class_id[i], :].values.tolist()
        sht[index_id + 2, col_nums-1].value = achieve.loc[class_id[i], :].values.tolist()

        sht[index_id + 1: index_id+3, col_nums - 1: col_nums + len(average.columns)-1].color = (242, 242, 242)          # 背景色
        sht[index_id + 1: index_id + 3, col_nums - 1: col_nums + len(average.columns)-1].api.Borders.LineStyle = 1      # 框线

# 学生总分、得分、达成度
def student_data_to_excel(sht, df_old, object_scores, col_nums, object_evaluation_out, everyclass_student, sheet_name, param_out, object_evaluation_achieve, df_test_weight, object_all_out):
    # 得分参数部分
    param_out.drop(index='小题分数', axis=0, inplace=True)
    param_out = param_out[param_out[0] != 0]

    # sht[1:3, col_nums: col_nums + len(param_out)].delete()

    sht[1, col_nums: col_nums + len(param_out)].api.Merge()
    sht[1, col_nums].value = str(sheet_name) + '得分'
    sht[2, col_nums].value = list(param_out.index)
    sht[3, col_nums].value = param_out.T.values.tolist()
    # 达成度参数部分

    # sht[1:2, col_nums + len(param_out): col_nums + len(param_out)*2].delete()

    sht[1, col_nums + len(param_out): col_nums + len(param_out)*2].api.Merge()
    sht[1, col_nums + len(param_out)].value = '达成度'
    sht[2, col_nums + len(param_out)].value = list(param_out.index)
    # 参数部分样式
    sht[1: 4 + len(df_test_weight.columns), col_nums-1: col_nums + len(param_out) * 2].api.Borders.LineStyle = 1     # 框线
    # sht[1: 4 + len(df_test_weight.columns), col_nums-1: col_nums + len(param_out) * 2].color = (242, 242, 242)       # 背景颜色

    # 学生成绩写入（总分、得分、达成度拼一起）
    df_out = pd.DataFrame(object_scores.T.sum())         # 总分
    object_all_out.append(df_out)
    all_data = pd.concat((df_out, object_evaluation_out), axis=1)
    all_data = pd.concat((all_data, object_evaluation_achieve), axis=1)
    all_data = all_data.loc[:, ~(all_data == 0).all(axis=0)]
    all_data = all_data.loc[:, ~(all_data == 'nan').all(axis=0)]
    for i in everyclass_student:
        index_id = find_it(df_old, i[0]) + 2
        sht[index_id, col_nums-1].value = all_data.loc[i].values.tolist()                                               # 写入数据
        sht[index_id: index_id + len(i), col_nums-1: col_nums + len(all_data.columns) - 1].api.Borders.LineStyle = 1    # 框线
    sht[1, col_nums - 1].value = '总分'
    return col_nums + len(all_data.columns) - 1, object_all_out

# 客观评价表格样式
def excel_style(sht, col_nums, end_col):
    info = sht.used_range
    nrows = info.last_cell.row          # 有效数据行数
    sht[1:nrows, col_nums - 1: end_col].api.HorizontalAlignment = -4108                     # 居中
    sht[1:nrows, col_nums- 1: end_col].api.Font.Size = 10                                   # 字体大小

# ‘达成度散点图’页帧数据写入
def final_data_to_excel(wb, all_target_data, student_id, df_index_values, df_weight, data_weight, df_test_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd):
    try:
        wb.sheets.add('达成度散点图', after=str(df_index_values[-1]))
        sht = wb.sheets['达成度散点图']
        target_name = df_weight.index       # 目标名称列表
        start_col = 1                       # 每个目标起始列标号
        start_idx = 1                       # 每个目标起始行标号
        #写入行索引
        sht[3, 0].value = '评价权重值'
        sht[5, 0].value = '课程目标权重值'
        index_id = student_id + ['平均值']
        sht[6, 0].options(transpose=True).value = index_id
        info = sht.used_range
        row = info.last_cell.row  # 获取’达成度散点图‘页帧当前有效数据行数
        sht[row + 1, 0:2].value = ['总人数', len(index_id) - 1]
        sht[row + 2, 0].options(transpose=True).value = ['单项目标超过平均值人数', '单项目标达成人数', '单项目标未达成人数', '所有目标超过平均值人数', '所有目标均达成人数', '所有目标未达成人数']

        for i in range(len(all_target_data)):
            # 处理各个目标数据
            single_target_data = all_target_data[i].loc[:, ~(all_target_data[i] == 'nan').all(axis=0)]
            all_target_data[i] = single_target_data
            # 拼索引dataframe并写入
            col_num = concat_index_df(sht, target_name[i], df_index_values, data_weight, single_target_data, df_test_weight, start_idx, start_col)
            # 写入每位同学成绩
            sht[6, start_col].value = single_target_data.values.tolist()
            # 写入达成度
            sht[6, start_col + len(single_target_data.columns)].options(transpose=True).value = final_target_achieve.loc[:, str(target_name[i])].values.tolist()
            # 写入均值
            sht[6, start_col + len(single_target_data.columns) + 1].options(transpose=True).value = [final_target_achieve.loc[:, str(target_name[i])].values[-1]] * (len(index_id)-1)
            # 添加统计数据
            sht[row + 2, start_col + len(single_target_data.columns)].options(transpose=True).value = static_final_data.loc[:, str(target_name[i])].values.tolist()
            sht[row + 2, start_col + len(single_target_data.columns) + 1].options(transpose=True).value = static_data_single_odd.loc[:, str(target_name[i])].values.tolist()
            start_col = start_col + col_num  # 更新列索引
        sht[row + 5, len(all_target_data[0].columns) + 2].options(transpose=True).value = static_data_all_odd.values.tolist()

        sht[1, start_col].value = '平均值'
        sht[6, start_col].options(transpose=True).value = [0.7] * (len(index_id) - 1)
        # 单元格样式
        nrows = 5 + len(index_id)
        sht[1: nrows, 0: start_col].api.Borders.LineStyle = 1               # 框线
        sht[1: nrows, 0: start_col].api.HorizontalAlignment = -4108         # 居中
        sht[1: nrows, 0: start_col].api.Font.Size = 10                      # 字体大小
        return start_col
    except Exception as error:
        logger.exception(error)

# 拼达成度散点图索引dataframe
def concat_index_df(sht, target_name, df_index_values, data_weight, single_target_data, df_test_weight, start_idx, start_col):
    try:
        # 拼索引dataframe
        col_num = len(single_target_data.columns) + 2
        df_index = pd.DataFrame(data=None, index=range(5), columns=range(col_num))
        df_index.loc[0, df_index.columns[-1]] = target_name                                         # 写入目标名称
        object_num = len(single_target_data.columns) - len(df_index_values)                         # 获取各个目标客观评价个数
        df_index.loc[1, object_num - 1] = '客观评价'
        df_index.loc[1, object_num: object_num+1] = df_index_values                                 # 写入主观评价名称
        df_index.loc[1, df_index.columns[-2]:] = ['达成度', '平均值']                                 # 写入达成度、平均值名称
        df_index.loc[2, object_num - 1: object_num + len(df_index_values) - 1] = data_weight.values.tolist()   # 写入评价权重值
        df_index.loc[3, 0: object_num - 1] = single_target_data.columns[:-len(df_index_values)]      # 写入客观评价名称
        object_weight = df_test_weight.loc[:, str(target_name)]                                      # 当前目标客观评价权重
        object_weight = object_weight[object_weight.values != 0]
        df_index.loc[4, 0:object_num - 1] = object_weight.values.tolist()                            # 写入客观评价权重
        # 写入excel
        sht[start_idx, start_col].value = df_index.values.tolist()
        # 合并单元格
        sht[1, start_col: start_col + col_num].api.Merge()
        sht[2, start_col: start_col + object_num].api.Merge()
        sht[3, start_col: start_col + object_num].api.Merge()
        # 背景色
        sht[3, start_col + object_num + 2: start_col + object_num + len(df_index_values) + 2].color = (242, 242, 242)
        sht[4: 6, start_col + object_num: start_col + object_num + len(df_index_values) + 2].color = (242, 242, 242)
        return col_num
    except Exception as error:
        logger.exception(error)

# 处理‘课程目标达成情况报告’页帧
def write_reporter(wb, sheet, df_test_weight, class_ach, class_index, df_index_values, data_weight):
    try:
        sht = wb.sheets['课程目标达成情况报告']
        # 找出指定班级
        df = pd.read_excel(sheet, sheet_name='课程目标达成情况报告', index_col=0, header=None)
        df = df.loc['班级', :]
        df = df.dropna(axis=0, how='all')
        df.index = range(len(df))
        class_id = df.loc[0]
        # 处理指定班级达成度
        data = pd.DataFrame(data=None, index=df_test_weight.columns)
        # print(class_ach)
        for i in range(len(df_test_weight.index)):
            single_ach = class_ach[i]
            index_type = class_ach[i].index.tolist()
            a = type(index_type[0])
            class_id = a(class_id)
            ach = single_ach.loc[class_id, :]
            data = pd.concat([data, ach], axis=1)
        data.columns = df_test_weight.index
        data = data.fillna(value=0)

        df = pd.read_excel(sheet, sheet_name='课程目标达成情况报告', header=None)
        df = df.dropna(axis=0, how='all')
        # 查找重复数据起始位置（班级达成度写入位置）
        repeat_data = df.loc[:len(df.index) - 2, 0].duplicated()
        repeat_index = repeat_data[repeat_data == True].index.tolist()
        row = repeat_index[0]
        sht[row + 1, 2].value = data.T.values.tolist()
        # 处理课程目标达成度情况
        weight = list()
        row2 = row + len(df_test_weight.index) + 3
        data = data.astype(float)
        achievement = (df_test_weight.mul(data.T, axis=1)).sum(axis=0)
        weight.append(achievement.values.tolist())
        sht[row2, 2].value = achievement.values.tolist()                     # 考核成绩法写入
        unobject_ach = class_ach[-len(df_index_values):]
        for i in range(len(unobject_ach)):
            ach = unobject_ach[i].loc[class_id, :]
            weight.append(ach.values.tolist())
            sht[row2+i, 2].value = ach.values.tolist()                    # 主观评价达成度写入
        class_weight = pd.DataFrame(data=weight)
        data_weight.index = range(len(data_weight))
        com_weight = (class_weight.mul(data_weight, axis=0)).sum(axis=0)
        sht[row2 + len(df_index_values), 2].value = com_weight.values.tolist()
        # # sht[row + 1, 0: len(df_test_weight.columns) + 2].api.Font.Size = 12                     # 字体大小
        # # sht[row + 2: row2 - 1, 0: len(df_test_weight.columns) + 2].api.Font.Size = 10           # 字体大小
    except Exception as error:
        logger.exception(error)

# 主观评价班级平均分、达成度写入
def unobject_class_data(sht, df_sheet, average, achieve, everyclass_student, class_id):
    for i in range(len(everyclass_student)):
        index_id = find_it(df_sheet, everyclass_student[i][-1])
        sht[index_id + 1, 2].value = average.loc[class_id[i], :].values.tolist()
        sht[index_id + 2, 2].value = achieve.loc[class_id[i], :].values.tolist()

# 散点图
def plt_to_excel(figure, sheet, target_name, start_col, start_idx):
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(sheet)
        sht = wb.sheets['达成度散点图']
        sht.pictures.add(figure, name=target_name, left=sht[start_idx, start_col+2].left, top=sht[start_idx, start_col+2].top)
        wb.save(sheet)
        wb.close()
        app.quit()
        app.kill()
    except Exception as error:
        logger.exception(error)