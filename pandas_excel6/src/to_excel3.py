from log.log import *
import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

# 定位表中things_needed字符所在的坐标
def find_it(dtx, things_needed):
    index_id = 'nan'
    for j in range(len(dtx.index)):
        data = dtx.loc[j]
        if things_needed in data.values:
            index_id = j
            data.index = range(len(data))
    return index_id

# 参数总分写入dataframe
def params_to_excel(df_old, object_param, df_test_weight):
    score_part = object_param.loc[df_test_weight.columns, :]
    object_out = (score_part * object_param.loc['小题分数']).T.sum()
    object_out['小题分数'] = object_param.loc['小题分数'].T.sum()
    col_old = df_old.shape[1]
    df_old.insert(col_old, '总分', object_out)
    df_old = df_old.T.reset_index()
    return df_old.T, object_out

# 分数总分写入dataframe
def scores_out(df_params,  object_scores, id, student_id):
    df_params.reset_index(inplace=True)
    df_params.index = df_params.loc[:, id]
    df_params.drop(columns=id, inplace=True)
    df_out = pd.DataFrame(object_scores.T.sum())
    for i in student_id:
        df_params.loc[i, len(df_params.columns)-1] = df_out.loc[i].values[0]
    return df_params

# 总分写入excel
def out_to_excel(df_old, object_all_params, df_test_weight, object_all_scores, student_id, id, sht):
    # 参数总分
    df_params_scores, object_out = params_to_excel(df_old, object_all_params, df_test_weight)
    # 得分总分
    df_out = scores_out(df_params_scores, object_all_scores, id, student_id)
    df_out.columns = range(len(df_out.columns))
    df_out = df_out.loc[:, df_out.columns[-1]]
    col_nums = len(df_old.columns)
    sht[1, col_nums].options(transpose=True).value = df_out.values.tolist()
    sht[1, col_nums].color = (242, 242, 242)      # 设置单元格背景色
    return col_nums + 1, object_out

# 将得分写入excel
def scores_to_excel(df_old, object_evaluation_out, everyclass_student, sheet_name, col_nums, object_out, sht):
    # 处理得分dataframe
    object_out = pd.DataFrame(object_out)
    object_out = object_out[object_out.values != 0]
    object_out = object_out.filter(axis=0, regex=str('目标'))
    sheet_head = pd.concat((object_out.T, pd.DataFrame([])), keys=[str(sheet_name)+'得分', ''], axis=1)
    sheet_head = sheet_head.T.reset_index()
    sht[1, col_nums].value = sheet_head.T.values.tolist()
    sht[1, col_nums:col_nums+len(object_out.index)].api.Merge()

    for i in everyclass_student:
        index_id = find_it(df_old, i[0])
        scores = object_evaluation_out.loc[i]
        df = scores.loc[:, ~(scores == 0).all(axis=0)]
        sht[index_id, col_nums].options(transpose=True).value = df.T.values.tolist()                        # 写入数据
        sht[index_id: index_id + len(i), col_nums: col_nums + len(df.columns)].api.Borders.LineStyle = 1    # 框线
    return col_nums+len(object_out.index)

# 将达成度写入excel
def achieve_to_excel(df_sheet, object_evaluation_achieve, everyclass_student, col_nums, sht):
    for i in everyclass_student:
        index_id = find_it(df_sheet, i[0])
        achieve = object_evaluation_achieve.loc[i]
        achieve = achieve.loc[:, ~(achieve == 'nan').all(axis=0)]
        sht[index_id, col_nums].options(transpose=True).value = achieve.T.values.tolist()
        sht[index_id: index_id + len(achieve), col_nums: col_nums + len(achieve.columns)].api.Borders.LineStyle = 1
        sht[1, col_nums:col_nums + len(achieve.columns)].api.Merge()
        sht[1, col_nums].value = '达成度'
        sht[2, col_nums].value = list(achieve.columns)
    return len(achieve.columns)

# 将班级平均分、达成度写入excel
def object_class_data(df_sheet, class_ach, class_avg, everyclass_student, col_nums, class_id, sht):
    for i in range(len(everyclass_student)):
        index_id = find_it(df_sheet, everyclass_student[i][-1])
        average = class_avg.loc[:, ~(class_avg == 0).all(axis=0)]
        achieve = class_ach.dropna(axis=1, how='all')
        sht[index_id + 1, col_nums].value = average.loc[class_id[i], :].values.tolist()
        sht[index_id + 2, col_nums].value = achieve.loc[class_id[i], :].values.tolist()
        sht[index_id + 1: index_id+3, col_nums -1: col_nums + len(average.columns)].color = (242, 242, 242)     # 背景色
        sht[index_id + 1: index_id+3, col_nums: col_nums + len(average.columns)].api.Borders.LineStyle = 1      # 框线

# 主观评价班级平均分、达成度写入
def unobject_class_data(sht, df_sheet, average, achieve, everyclass_student, class_id):
    for i in range(len(everyclass_student)):
        index_id = find_it(df_sheet, everyclass_student[i][-1])
        sht[index_id + 1, 2].value = average.loc[class_id[i], :].values.tolist()
        sht[index_id + 2, 2].value = achieve.loc[class_id[i], :].values.tolist()

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
        df_index.loc[1, df_index.columns[-2]:] = ['达成度', '平均值']                                                  # 写入达成度、平均值名称
        df_index.loc[2, object_num - 1: object_num + len(df_index_values) - 1] = data_weight.values.tolist()   # 写入评价权重值
        df_index.loc[3, 0: object_num - 1] = single_target_data.columns[:-len(df_index_values)]      # 写入客观评价名称
        object_weight = df_test_weight.loc[:, str(target_name)]                                      # 当前目标客观评价权重
        object_weight = object_weight[object_weight.values != 0]
        df_index.loc[4, 0:object_num - 1] = object_weight.values.tolist()                            # 写入客观评价权重
        # 写入excel
        sht[start_idx, start_col].value = df_index.values.tolist()
        # 合并单元格
        sht[1, start_col: start_col + col_num].api.Merge()         # 目标名称单元格
        sht[2, start_col: start_col + object_num].api.Merge()         # 目标名称单元格

        return col_num
    except Exception as error:
        logger.exception(error)


# ‘达成度散点图’页帧数据写入
def final_data_to_excel(wb, all_target_data, student_id, df_index_values, df_weight, data_weight, df_test_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd):
    wb.sheets.add('达成度散点图', after=str(df_index_values[-1]))
    sht = wb.sheets['达成度散点图']
    target_name = df_weight.index     # 目标名称列表
    start_col = 1                       # 每个目标起始列标号
    start_idx = 1                       # 每个目标起始行标号
    #写入行索引
    sht[3, 0].value = '评价权重值'
    sht[5, 0].value = '课程目标权重值'
    student_id.append('平均值')
    sht[6, 0].options(transpose=True).value = student_id
    info = sht.used_range
    row = info.last_cell.row  # 获取’达成度散点图‘页帧当前有效数据行数
    sht[row + 1, 0:2].value = ['总人数', len(student_id) - 1]
    sht[row + 2, 0].options(transpose=True).value = ['单项目标超过平均值人数', '单项目标达成人数', '单项目标未达成人数', '所有目标超过平均值人数', '所有目标均达成人数', '所有目标未达成人数']
    for i in range(len(all_target_data)):
        # 处理各个目标数据
        single_target_data = all_target_data[i].loc[:, ~(all_target_data[i] == 'nan').all(axis=0)]
        # 拼索引dataframe并写入
        col_num = concat_index_df(sht, target_name[i], df_index_values, data_weight, single_target_data, df_test_weight, start_idx, start_col)
        # 写入每位同学成绩
        sht[6, start_col].value = single_target_data.values.tolist()
        # 写入达成度
        sht[6, start_col + len(single_target_data.columns)].options(transpose=True).value = final_target_achieve.loc[:, str(target_name[i])].values.tolist()
        # 写入均值
        sht[6, start_col + len(single_target_data.columns) + 1].options(transpose=True).value = [final_target_achieve.loc[:, str(target_name[i])].values[-1]] * (len(student_id)-1)

        # 添加统计数据
        sht[row + 2, start_col + len(single_target_data.columns)].options(transpose=True).value = static_final_data.loc[:, str(target_name[i])].values.tolist()
        sht[row + 2, start_col + len(single_target_data.columns) + 1].options(transpose=True).value = static_data_single_odd.loc[:, str(target_name[i])].values.tolist()

        start_col = start_col + col_num  # 更新列索引
    sht[row + 5, len(all_target_data[0].columns) + 1].options(transpose=True).value = static_data_all_odd.values.tolist()

    sht[1, start_col].value = '平均值'
    sht[6, start_col].options(transpose=True).value = [0.7] * (len(student_id) - 1)
    # 单元格样式
    nrows = 5 + len(student_id)
    sht[1: nrows, 0: start_col].api.Borders.LineStyle = 1               # 框线
    sht[1: nrows, 0: start_col].api.HorizontalAlignment = -4108         # 居中
    sht[1: nrows, 0: start_col].api.Font.Size = 10                      # 字体大小
    return start_col

# 处理客观评价数据写入
def write_out_achieve(wb, object_all_params, object_all_scores, object_evaluation_achieve, object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k, student_id, everyclass_student, class_id, df_index_values):
    try:
        sheet_name = df_test_weight.index
        for i in range(len(sheet_name)):
            if i != k:
                df_old = pd.DataFrame(pd.read_excel(sheet, sheet_name=str(sheet_name[i]), header=1, index_col=1))
                id = 0
            else:
                df_old = pd.DataFrame(pd.read_excel(sheet, sheet_name=str(sheet_name[i]), header=1, index_col=2))
                id = 1
            df_sheet = pd.read_excel(sheet, sheet_name=str(sheet_name[i]), header=None)
            sht = wb.sheets[str(sheet_name[i])]
            # 将总分写入excel
            col_nums, object_out = out_to_excel(df_old, object_all_params[i], df_test_weight, object_all_scores[i], student_id, id, sht)
            # 将客观评价班级平均分、达成度写入excel
            object_class_data(df_sheet, class_ach[i], class_avg[i], everyclass_student, col_nums, class_id, sht)
            # 将得分写入excel
            col_nums = scores_to_excel(df_sheet, object_evaluation_out[i], everyclass_student, sheet_name[i], col_nums, object_out, sht)
            # 将达成度写入excel
            achieve_nums = achieve_to_excel(df_sheet, object_evaluation_achieve[i], everyclass_student, col_nums, sht)
            # 将写入数据部分设置居中和框线
            excel_style(sht, len(df_sheet.columns), achieve_nums, df_test_weight)
        # 将主观评价班级平均分、达成度写入表格
        average = class_avg[-len(df_index_values):]
        achieve = class_ach[-len(df_index_values):]
        for i in range(len(df_index_values)):
            sht = wb.sheets[df_index_values[i]]
            df_sheet = pd.read_excel(sheet, sheet_name=df_index_values[i], header=None)
            unobject_class_data(sht, df_sheet, average[i], achieve[i], everyclass_student, class_id)
    except Exception as error:
        logger.exception(error)

# 处理‘课程目标达成情况报告’页帧
def write_reporter(wb, sheet, df_test_weight, class_ach, class_index, df_index_values, data_weight):
    try:
        sht = wb.sheets['课程目标达成情况报告']
        # 找出指定班级
        rng = sht[4, :]
        for cell in rng:
            if cell.value in class_index:
                class_id = cell.value
        # 处理指定班级达成度
        data = pd.DataFrame(data=None, index=df_test_weight.columns)
        for i in range(len(df_test_weight.index)):
            single_ach = class_ach[i]
            ach = single_ach.loc[class_id, :]
            data = pd.concat([data, ach], axis=1)
        data.columns = df_test_weight.index
        data = data.fillna(value=0)
        # 插入所需行数
        rng = sht[:, 0]
        for cell in rng:
            if cell.value == '权重和':
                row = int(cell.address[3:])
        for i in range(len(df_test_weight.index) + 1):
            sht.api.Rows(row + 1).Insert()
        sht[row, 1].value = data.T
        # 合并单元格
        for i in range(len(df_test_weight.index)+1):
            sht[row + i, 0:2].api.Merge()
        sht[row, 0].value = '考核方式'
        # 处理课程目标达成度情况
        weight = list()
        row2 = row + len(df_test_weight.index)+2
        data = data.astype(float)
        achievement = (df_test_weight.mul(data.T, axis=1)).sum(axis=0)
        weight.append(achievement.values.tolist())
        sht[row2, 2].value = achievement.values.tolist()                     # 达成度（考核成绩法）写入
        unobject_ach = class_ach[-len(df_index_values):]
        for i in range(len(unobject_ach)):
            ach = unobject_ach[i].loc[class_id, :]
            weight.append(ach.values.tolist())
            sht[row2+1+i, 2].value = ach.values.tolist()                    # 主观评价达成度写入
        class_weight = pd.DataFrame(data=weight)
        data_weight.index = range(len(data_weight))
        com_weight = (class_weight.mul(data_weight, axis=0)).sum(axis=0)
        sht[row2 + len(df_index_values)+1, 2].value = com_weight.values.tolist()
        sht[row: row2 - 1, 0: len(df_test_weight.columns) + 2].api.Font.Size = 10  # 字体大小
    except Exception as error:
        logger.exception(error)

# 表格样式
def excel_style(sht, col_nums, achieve_nums, df_test_weight):
    info = sht.used_range
    nrows = info.last_cell.row
    ncols = achieve_nums * 2 + 1
    # sht[1:nrows, col_nums:col_nums + ncols].api.Borders.LineStyle = 1                           # 框线
    sht[1:nrows, col_nums].api.Borders.LineStyle = 1                           # 总分列框线
    sht[1:nrows, col_nums:col_nums + ncols].api.HorizontalAlignment = -4108                     # 居中
    sht[1:nrows, col_nums:col_nums + ncols].api.Font.Size = 10                                  # 字体大小
    a = 6 + len(df_test_weight.columns)
    sht[1: a, col_nums - 1:col_nums + ncols].color = (242, 242, 242)             # 背景颜色
    sht[1: a, col_nums - 1:col_nums + ncols].api.Borders.LineStyle = 1          # 表头框线

# 将散点图写入excel
def plt_to_excel(figure, sheet, target_name, start_col, start_idx):
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(sheet)
        sht = wb.sheets['达成度散点图']

        sht.pictures.add(figure, name=target_name, left=sht[start_idx, start_col+2].left, top=sht[start_idx, start_col + 2].top)

        wb.save(sheet)
        wb.close()
        app.quit()
        app.kill()
    except Exception as error:
        logger.exception(error)

def write_excel(object_all_params, object_all_scores, object_evaluation_achieve,
                                            object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k,
                                            student_id, everyclass_student, class_id, df_index_values, all_target_data, df_weight, data_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(sheet)
    app.display_alerts=False
    # 将评价页帧计算数据写入excel
    write_out_achieve(wb, object_all_params, object_all_scores, object_evaluation_achieve,
                                            object_evaluation_out, df_test_weight, class_ach, class_avg, sheet, k,
                                            student_id, everyclass_student, class_id, df_index_values)
    # 添加‘达成度散点图’页帧
    start_col = final_data_to_excel(wb, all_target_data, student_id, df_index_values, df_weight, data_weight, df_test_weight, final_target_achieve, static_final_data, static_data_single_odd, static_data_all_odd)

    # 处理‘课程目标达成情况报告’页帧
    write_reporter(wb, sheet, df_test_weight, class_ach, class_id, df_index_values, data_weight)

    new_path = '../data_new/原表修改2.xls'
    wb.save(new_path)
    wb.close()
    app.quit()
    app.kill()
    return start_col, new_path

