from src.main import *
from src import to_function
from src.function import *
from src.get_info import sheet
from log.log import logger

name_sheet = df_weight.columns
def get_avg(option):
    lst_excel_class_score = list()
    try:
        for i in df_weight.columns:
            df_sheet = pd.read_excel(sheet, sheet_name=str(i), header=None)
            index_axis, columns_axis = find_it(df_sheet, '学生姓名')
            df_sheet.columns = df_sheet.loc[index_axis, :]
            df_sheet.index = df_sheet.loc[:, '学生姓名']
            df_sheet.drop(columns=['学生姓名', '学生学号'], axis=1, inplace=True)
            df_sheet = df_sheet.loc[option, :].dropna(axis=1, how='all').applymap("{0:.03f}".format)
            df_sheet.index = class_id
            columns = df_weight[df_weight[str(i)] != 0].index
            df_sheet.columns = columns
            lst_excel_class_score.append(df_sheet)
        return lst_excel_class_score
    except Exception as error:
        logger.exception('出现异常', error)

# 各表格平均值
def get_class_avg():
    # 提取excel平均值
    try:
        lst_excel_class_score = get_avg('平均分')
        return lst_excel_class_score, class_avg, name_sheet
    except Exception as error:
        logger.exception('出现异常', error)

# 各表格平均值
def get_class_ach():
    try:
        # 提取excel平均值
        lst_excel_class_score = get_avg('达成度')
        # print(class_ach)
        # 提取计算数据
        class_ach, class_avg = to_function.count_class_avg_ach(object_evaluation_out, object_all_param, unobject_scores,
                                                               unobject_param, df_weight, class_id, everyclass_student)
        return lst_excel_class_score, class_ach, name_sheet
    except Exception as error:
        logger.exception('出现异常', error)

def get_every_scores_achieve():
    try:
        # 计算数据
        object_evaluation_achieve, object_evaluation_out, object_all_param = to_function.count_object_target(
            object_all_scores, object_param, df_weight)
        lst_count_data = list()
        for i in range(len(object_evaluation_out)):
            score_achieve = pd.merge(object_evaluation_out[i], object_evaluation_achieve[i], left_index=True,
                                     right_index=True, suffixes=('_得分', '_达成度'))
            score_achieve = score_achieve.loc[:, (score_achieve != 0).any(axis=0)]
            score_achieve = score_achieve.loc[:, (score_achieve != 'nan').all(axis=0)]
            lst_count_data.append(score_achieve)
    except Exception as error:
        logger.exception('出现异常', error)

    try:
        # excel数据
        lst_excel_data = list()
        for i in df_test_weight.index:
            k = (df_test_weight.loc[str(i), :] != 0).sum()
            df_sheet = pd.read_excel(sheet, sheet_name=str(i))
            index_axis, columns_axis = find_it(df_sheet, '学生学号')
            df_sheet.columns = df_sheet.loc[index_axis, :]
            df_sheet.index = df_sheet.loc[:, '学生学号']
            df_sheet.drop(columns='学生学号', axis=1, inplace=True)
            df_sheet.columns = range(len(df_sheet.columns))
            scores = df_sheet.loc[student_id, len(df_sheet.columns)-(k*2):len(df_sheet.columns)-k-1]
            columns = df_test_weight.loc[str(i)]
            columns = columns[columns.values != 0]
            scores.columns = columns.index
            achieve = df_sheet.loc[student_id, len(df_sheet.columns) - k:]
            achieve.columns = columns.index
            df_sheet = pd.merge(scores, achieve, left_index=True, right_index=True, suffixes=('_得分', '_达成度'))
            lst_excel_data.append(df_sheet)
        return lst_count_data, lst_excel_data, name_sheet
    except Exception as error:
        logger.exception('出现异常', error)

# 达成度和均值对比
def get_final_achieve():
    try:
        # 计算出的达成度和均值
        count_achieve = to_function.count_final_achieve(all_target_data, df_weight, student_id)
    except Exception as error:
        logger.exception('出现异常', error)

    try:
        # excel数据
        excel_achieve = pd.read_excel(sheet, sheet_name='达成度散点图', header=2, index_col=0)
        excel_achieve = excel_achieve.filter(regex='达成度')
        excel_achieve.columns = df_weight.index
        excel_achieve = excel_achieve.loc[1:'平均值', :]
        excel_achieve.index = count_achieve.index
        excel_achieve = excel_achieve.applymap("{0:.03f}".format)
        return count_achieve, excel_achieve
    except Exception as error:
        logger.exception('出现异常', error)

# 统计部分
def get_static_data():
    try:
        # 计算数据
        achieve, static_final_data, static_data_single_odd, static_data_all_odd = function.static_data(target_achieve, student_id, df_weight)
    except Exception as error:
        logger.exception('出现异常', error)

    # excel数据
    try:
        df_sheet = pd.read_excel(sheet, sheet_name='达成度散点图', header=2, index_col=0)
        df_sheet = df_sheet.loc['单项目标超过平均值人数':, :]
        df_sheet = df_sheet.dropna(axis=1, how='all')
        count_data = df_sheet.filter(regex=str('达成度'))
        count_data.columns = df_weight.index

        count_data_odd = df_sheet.filter(regex=str('平均值'))
        count_data_single_odd = count_data_odd.loc['单项目标超过平均值人数':'单项目标未达成人数', :]
        count_data_all_odd = count_data_odd.loc['所有目标超过平均值人数':, :].dropna(axis=1, how='all')
        count_data_all_odd = count_data_all_odd.loc[:, '平均值']

        count_data_single_odd = count_data_single_odd.applymap(lambda x: format(x, '.2%'))
        count_data_all_odd = count_data_all_odd.map(lambda x: format(x, '.2%'))
        count_data_single_odd.columns = df_weight.index
        return static_final_data, static_data_single_odd, static_data_all_odd, count_data, count_data_single_odd, count_data_all_odd
    except Exception as error:
        logger.exception('出现异常', error)
