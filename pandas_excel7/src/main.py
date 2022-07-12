from log.log import *
from src import check_data
import os
from src import get_info

sheet_name = get_info.get_excel_name()
for file in sheet_name:
    try:
        file_name = os.path.split(file)
        file_name = os.path.splitext(file_name[1])
        get_info.log_add(file_name[0])
        check_data.check_cover(file)  # 检查封皮数据和报告页帧是否一致
        # 提取课程目标达成情况报告信息

        df_weight, df_test_weight, df_index_values, data_weight = get_info.get_weight(file)
        logger.info('\n 权重矩阵：{name}\n{}', df_weight, name='df_weight')
        # 达成度
        final_target_achieve, new_path, student_id, k, object_all_out = get_info.operate_sheet(df_test_weight, file, df_index_values, df_weight, data_weight)
        print('程序运行完成')
    except Exception:
        print(file_name[0], '表格出现异常')