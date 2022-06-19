# import pytest
# from log.log import logger
# from src.main import new_path, df_weight, df_test_weight, student_id, k, df_index_values
# import pandas as pd
#
# @pytest.fixture(autouse=True)
# def set_up():
#     yield
#     print("\n测试结束")
#
# # 页帧个数是否满足需求（客观评价+主观评价+达成度散点图）
# def test_class_avg():
#     print('\n正在执行测试模块-----页帧个数是否满足需求')
#     logger.debug('\n正在执行测试模块-----页帧个数是否满足需求')
#     # 所需页帧
#     sheet_lst = list(df_weight.columns)
#     sheet_lst = sheet_lst + ['封皮', '课程目标达成情况报告', '达成度散点图']
#
#     # 新excel页帧
#     df = pd.read_excel(new_path, sheet_name=None)
#     new_excel_sheet = list(df)
#
#     sheet_lst.sort()
#     new_excel_sheet.sort()
#     if sheet_lst != new_excel_sheet:
#         if len(sheet_lst) >= len(new_excel_sheet):
#             no_sheet = [x for x in sheet_lst if x not in new_excel_sheet]
#             print('缺少页帧：', no_sheet)
#             logger.error('缺少页帧：{no_sheet}', no_sheet=no_sheet)
#         if len(sheet_lst) <= len(new_excel_sheet):
#             no_sheet = [x for x in new_excel_sheet if x not in sheet_lst]
#             print('多出页帧：', no_sheet)
#             logger.error('多出页帧：{no_sheet}', no_sheet=no_sheet)
#     else:
#         logger.info('页帧个数满足需求')
#     assert sheet_lst == new_excel_sheet
#
# # '课程目标达成情况报告'页帧：班级达成度数据是否写入、是否规范
# def test_report():
#     print('\n正在执行测试模块-----\'课程目标达成情况报告\'页帧是否满足需求')
#     logger.debug('\n正在执行测试模块-----\'课程目标达成情况报告\'页帧是否满足需求')
#     df = pd.read_excel(new_path, sheet_name='课程目标达成情况报告', index_col=0, header=None)
#     df = df.loc['权重和':, :]
#     df.index = map(lambda x: str(x).replace('\n', ' '), [x for x in df.index])
#     # 所有指标是否写入
#     data_index = list(df_weight.columns)
#     other_data_index = ['考核成绩法', '综合加权']
#     data_index.extend(other_data_index)
#     if set(data_index) < set(df.index):
#         logger.info('班级达成度数据均已写入')
#         # 目标个数不一致的索引
#         class_data = df.loc[data_index, :]
#         class_data.dropna(axis=1, how='any', inplace=True)
#         target_num = len(df_weight.index)
#         no_data = list(class_data[class_data.count(axis=1) != target_num].index)
#         # 写入数据是否在[0,1]
#         result = class_data[(class_data <= 1) & (class_data >= 0)]
#         error_data = list(result[pd.isnull(result).any(axis=1)].index)
#     else:
#         no_evaluate = [item for item in list(df_weight.columns) if item not in set(df.index)]
#         logger.error('{error_name}评价方式达成度未写入', error_name=no_evaluate)
#     # 写入目标个数是否一致
#     if no_data:
#         logger.error('目标个数不一致的索引：{error_num}', error_num=no_data)
#     else:
#         logger.info('目标个数一致')
#     # 写入数据是否规范
#     if error_data:
#         logger.error('{error_num} 写入数据不合法', error_num=error_data)
#     else:
#         logger.info('班级达成度数据写入合法')
#
#     pytest.assume(set(data_index) < set(df.index))
#     pytest.assume(len(no_data) == 0)
#     pytest.assume(len(error_data) == 0)
#
# # 客观评价页帧：得分、达成度、班级平均分和达成度是否写入规范
# def test_object_sheet():
#     print('\n正在执行测试模块-----客观评价页帧是否满足需求')
#     logger.debug('\n正在执行测试模块-----客观评鉴页帧是否满足需求')
#     sheet_name = df_test_weight.index.tolist()
#     for i in range(len(sheet_name)):
#         if i != k:
#             df = pd.read_excel(new_path, sheet_name=str(sheet_name[i]), header=[1, 2], index_col=0)
#         else:
#             df = pd.read_excel(new_path, sheet_name=str(sheet_name[i]), header=[1, 2], index_col=[0, 1])
#             df.index = df.index.droplevel()
#         df_columns = df.columns
#         need_data = df_test_weight.loc[str(sheet_name[i]), :]
#         need_data = need_data[need_data != 0].index.tolist()   # 获取每个客观评价所需写的目标
#         # 查看总分是否写入
#
#         # 查看得分列是否写入
#         for j in need_data:
#             if (str(sheet_name[i])+'得分', str(j)) not in df_columns:
#                 print(str(sheet_name[i]) + str(j) + '得分'+"  数据写入异常")
#             else:
#                 # 判断是否全部学生都写入
#                 score_data = df.loc[:, (str(sheet_name[i])+'得分', str(j))]
#                 score_data = score_data.dropna(axis=0, how='any')
#                 idx = list(set(score_data.index.tolist()))
#                 if not set(student_id) < set(idx):
#                     print(str(sheet_name[i]) + str(j) + '得分'+"  数据未写入全部学生数据")
#         # 查看达成度列是否写入
#         for j in need_data:
#             if ('达成度', str(j)) not in df_columns:
#                 print(str(sheet_name[i]) + str(j) + '达成度'+"  数据写入异常")
#             else:
#                 # 判断是否全部学生都写入
#                 score_data = df.loc[:, ('达成度', str(j))]
#                 score_data = score_data.dropna(axis=0, how='any')
#                 idx = score_data.index
#                 if not set(student_id) == set(idx):
#                     print(str(sheet_name[i]) + str(j) + '达成度'+"  数据未写入全部学生数据")
#         # 查看班级平均分、达成度是否写入
#
# # 主观评价页帧：班级平均分和达成度是否写入规范
# # '达成度散点图'页帧：课程目标个数是否一致、所有学生数据是否写入、达成度散点图是否写入、统计数据是否写入.....
#
#
# if __name__ == '__main__':
#     pytest.main(['-vs'])