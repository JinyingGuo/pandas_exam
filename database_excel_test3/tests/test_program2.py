import pytest
from tests import test_get_data2
from log.log import logger
# pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

@pytest.fixture(autouse=True)
def set_up():
    yield
    print("\n测试结束")

def test_class_avg():
    print('\n正在执行测试模块-----班级平均分')

    lst_excel_average, class_avg, name_sheet = test_get_data2.get_class_avg()

    # 计算的数据
    count_average = list()
    for i in class_avg:
        average = i.dropna(axis=1, how='all')
        count_average.append(average)

    if(len(lst_excel_average) == len(count_average)):
        for i in range(len(lst_excel_average)):
            excel_data = lst_excel_average[i].astype(float)
            result = excel_data == count_average[i]
            if False in result.values:
                print('计算数据：', count_average[i])
                print('excel数据：', excel_data)
                for j in count_average[i].columns:
                    logger.info(name_sheet[i] + "  页帧" + str(j) + "数据不一致的班级：{detail}", detail=(result[result == False].index.tolist()))
            else:
                logger.info(name_sheet[i] + "  班级平均分测试通过")

def test_class_ach():
    print('\n正在执行测试模块-----班级达成度')
    lst_achieve, class_ach, name_sheet = test_get_data2.get_class_ach()
    count_achieve = list()
    for i in class_ach:
        achieve = i.dropna(axis=1, how='all')
        count_achieve.append(achieve)
    if (len(lst_achieve) == len(count_achieve)):
        for i in range(len(count_achieve)):
            excel_data = lst_achieve[i].astype(float)
            result = excel_data == count_achieve[i]
            if False in result.values:
                print('计算数据：', count_achieve)
                print('excel数据：', lst_achieve)
                for j in count_achieve.columns:
                    logger.info(name_sheet[i] + str(j) + "   数据不一致的班级：{detail}", detail=(result[result == False].index.tolist()))
            else:
                logger.info(name_sheet[i] + "  班级达成度测试通过")

def test_target_scores():
    print('\n正在执行测试模块-----各目标得分和达成度')
    lst_count_data, lst_excel_data, name_sheet = test_get_data2.get_every_scores_achieve()
    if (len(lst_count_data) == len(lst_excel_data)):
        for i in range(len(lst_excel_data)):
            count_data = lst_count_data[i].astype(float)
            excel_data = lst_excel_data[i].astype(float)
            count_data = count_data.applymap("{0:.03f}".format)
            excel_data = excel_data.applymap("{0:.03f}".format)

            result = count_data == excel_data

            if False in result.values:
                print('计算数据：', count_data)
                print('excel数据：', excel_data)
                for j in count_data.columns:
                    logger.info(name_sheet[i] + "页帧    " + str(j) + "  数据不一致的学号：{detail}", detail=(result[result[str(i)] == False].index.tolist()))
            else:
                logger.info(name_sheet[i] + "  各目标得分、达成度测试通过")

def test_final_achieve():
    print('\n正在执行测试模块-----最终达成度')
    count_achieve, excel_achieve = test_get_data2.get_final_achieve()
    excel_achieve = excel_achieve.astype(float)

    result = count_achieve == excel_achieve
    if False in result.values:
        print('计算数据：', count_achieve)
        print('excel数据：', excel_achieve)
        for i in result.columns:
            logger.info(str(i) + "达成度数据不一致的学号：{detail}", detail=(result[result[str(i)] == False].index.tolist()))
    else:
        logger.info("达成度测试通过")

def test_static():
    print('\n正在执行测试模块-----统计部分')
    static_final_data, static_data_single_odd, static_data_all_odd, count_data, count_data_single_odd, count_data_all_odd = test_get_data2.get_static_data()
    static_final_data = static_final_data.fillna(value=0)
    count_data = count_data.fillna(value=0)

    result = static_final_data == count_data
    if False in result.values:
        print('计算数据：', count_data)
        print('excel数据：', static_final_data)
        for i in result.columns:
            logger.info(str(i) + "统计数据不一致的项：{detail}", detail=(result[result[str(i)] == False].index.tolist()))
    else:
        print("统计数据测试通过")

    result1 = static_data_single_odd == count_data_single_odd
    if False in result1.values:
        print('计算数据：', count_data_single_odd)
        print('excel数据：', static_data_single_odd)
        for i in result1.columns:
            logger.info(str(i) + "统计数据不一致的项：{detail}", detail=(result1[result1[str(i)] == False].index.tolist()))
    else:
        logger.info("统计数据测试通过")

    result2 = static_data_all_odd == count_data_all_odd
    if False in result2.values:
        print('计算数据：', count_data_all_odd)
        print('excel数据：', static_data_all_odd)
        logger.info("统计数据不一致的项：{detail}", detail=(result2[result2 == False].index.tolist()))
    else:
        logger.info("统计数据测试通过")

if __name__ == '__main__':
    pytest.main(['-vs'])