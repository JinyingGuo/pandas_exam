import pytest
from tests import test_get_data
import random

# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

@pytest.fixture(autouse=True)
def set_up():
    yield
    print("\n测试结束")

def test_object_class_avg():
    print('\n正在执行测试模块-----班级平均分')
    lst_average, class_avg, name_sheet = test_get_data.get_class_avg()

    # 随机取一个表
    sheet = random.choice(range(len(lst_average)))

    # 计算的数据
    class_average = class_avg[sheet].dropna(axis=1, how='all')
    # excel数据
    class_avg_data = lst_average[sheet]
    class_avg_data.columns = class_average.columns
    class_avg_data = class_avg_data.applymap("{0:.03f}".format)
    class_avg_data = class_avg_data.astype(float)
    print("随机选择的页帧：", name_sheet[sheet])
    print(class_average)
    print(class_avg_data)
    result = class_average == class_avg_data
    if False in result.values:
        print(result)
        print("数据不一致的索引：", result[result == 'False'].index.tolist())
    assert class_average.equals(class_avg_data)

def test_object_class_ach():
    print('\n正在执行测试模块-----班级达成度')
    lst_achieve, class_ach, name_sheet = test_get_data.get_class_avg()

    # 随机取一个表
    sheet = random.choice(range(len(lst_achieve)))

    # 计算的数据
    class_achieve = class_ach[sheet].dropna(axis=1, how='all')

    # excel数据
    class_ach_data = lst_achieve[sheet]
    class_ach_data.columns = class_achieve.columns
    class_ach_data = class_ach_data.applymap("{0:.03f}".format)
    class_ach_data = class_ach_data.astype(float)
    print("随机选择的页帧：", name_sheet[sheet])
    result = class_achieve == class_ach_data
    if False in result.values:
        print(result)
        print("数据不一致的索引：", result[result == 'False'].index.tolist())
    assert class_achieve.equals(class_ach_data)

def test_target_scores():
    print('\n正在执行测试模块-----各目标得分和达成度')
    lst_scores, lst_excel_scores, name_sheet = test_get_data.get_scores()

    # 随机取一个表
    sheet = random.choice(range(len(lst_scores)))
    print("随机选择的页帧：", name_sheet[sheet])
    # 计算的各目标得分
    test_score = lst_scores[sheet]
    test_score = test_score.loc[:, (test_score != 0).any(axis=0)]
    test_score = test_score.loc[:, (test_score != 'nan').all(axis=0)]
    test_score = test_score.astype(float)
    test_score = test_score.applymap("{0:.02f}".format)

    # excel数据
    df_excel_data = lst_excel_scores[sheet].dropna(axis=0, how='all')
    df_excel_data = df_excel_data.dropna(axis=1, how='all')
    df_excel_data.columns = test_score.columns
    df_excel_data = df_excel_data.applymap("{0:.02f}".format)
    result = test_score == df_excel_data
    if False in result.values:
        print(result)
        print("数据不一致的索引：", result[result == 'False'].index.tolist())
    assert test_score.equals(df_excel_data)

def test_final_achieve():
    print('\n正在执行测试模块-----最终达成度')
    count_achieve, excel_achieve = test_get_data.get_final_achieve()
    excel_achieve = excel_achieve.applymap("{0:.03f}".format)
    excel_achieve = excel_achieve.astype(float)
    result = count_achieve == excel_achieve
    if False in result.values:
        print(result)
        print("数据不一致的索引：", result[result == 'False'].index.tolist())
    assert count_achieve.equals(excel_achieve)

def test_static():
    print('\n正在执行测试模块-----统计部分')
    static_count_data, static_excel_data = test_get_data.static_data()
    static_count_data = static_count_data.fillna(value=0)
    static_excel_data = static_excel_data.fillna(value=0)
    result = static_count_data == static_excel_data
    if False in result.values:
        print(result)
        print("数据不一致的索引：", result[result == 'False'].index.tolist())
    assert static_count_data.equals(static_excel_data)

def test_all_data():
    print('\n正在执行测试模块-----达成度散点图页帧对比')
    count_data, excel_data = test_get_data.all_data_contrast()
    result = count_data == excel_data
    if False in result.values:
        print(result)
        print('数据不一致的索引：', result[result == 'False'].index.tolist())
    assert count_data.equals(excel_data)


# if __name__ == '__main__':
#     pytest.main(['-vs'])