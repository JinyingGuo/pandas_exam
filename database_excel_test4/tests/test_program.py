import pytest
from tests import test_get_data
from log.log import logger

@pytest.fixture(autouse=True)
def set_up():
    yield
    print("\n测试结束")

def test_class_avg():
    print('\n正在执行测试模块-----班级平均分')
    logger.debug('\n正在执行测试模块-----班级平均分')
    lst_excel_average, class_avg, name_sheet = test_get_data.get_class_avg()

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
                logger.info('\n计算数据：\n{detail}', detail=count_average[i])
                logger.info('\nexcel数据：\n{detail}', detail=excel_data)
                for j in count_average[i].columns:
                    detail = result[result[str(j)] == False].index.tolist()
                    if len(detail) != 0:
                        logger.info(name_sheet[i] + "页帧  " + str(j) + "  班级平均分数据不一致的班级：{detail}", detail=detail)
            else:
                logger.debug(name_sheet[i] + "  班级平均分测试通过")
    # pytest.assume(lst_excel_average, count_average)

def test_class_ach():
    print('\n正在执行测试模块-----班级达成度')
    logger.debug('\n正在执行测试模块-----班级达成度')
    lst_achieve, class_ach, name_sheet = test_get_data.get_class_ach()
    count_achieve = list()
    for i in class_ach:
        achieve = i.dropna(axis=1, how='all')
        count_achieve.append(achieve)
    if (len(lst_achieve) == len(count_achieve)):
        for i in range(len(count_achieve)):
            excel_data = lst_achieve[i].astype(float)
            result = excel_data == count_achieve[i]
            if False in result.values:
                logger.info('\n计算数据：\n{detail}', detail=count_achieve[i])
                logger.info('\nexcel数据：\n{detail}', detail=lst_achieve[i])
                for j in count_achieve[i].columns:
                    detail = result[result[str(j)] == False].index.tolist()
                    if len(detail) != 0:
                        logger.info(name_sheet[i] + '页帧  ' + str(j) + "   班级达成度  数据不一致的班级：{detail}", detail=detail)
            else:
                logger.debug(name_sheet[i] + "  班级达成度测试通过")
    # pytest.assume(lst_achieve, count_achieve)

def test_target_scores():
    print('\n正在执行测试模块-----各目标得分和达成度')
    logger.debug('\n正在执行测试模块-----各目标得分和达成度')
    lst_count_data, lst_excel_data, name_sheet = test_get_data.get_every_scores_achieve()
    if (len(lst_count_data) == len(lst_excel_data)):
        for i in range(len(lst_excel_data)):
            count_data = lst_count_data[i].astype(float)
            excel_data = lst_excel_data[i].astype(float)
            count_data = count_data.applymap("{0:.03f}".format)
            excel_data = excel_data.applymap("{0:.03f}".format)

            result = count_data == excel_data

            if False in result.values:
                logger.info('\n计算数据：\n{detail}', detail=count_data)
                logger.info('\nexcel数据：\n{detail}', detail=excel_data)
                for j in excel_data.columns:
                    detail = result[result[str(j)] == False].index.tolist()
                    if len(detail) != 0:
                        logger.info(name_sheet[i] + "页帧  " + str(j) + "  数据不一致的学号：{detail}", detail=detail)
            else:
                logger.debug(name_sheet[i] + "  各目标得分、达成度测试通过")
    # pytest.assume(lst_count_data, lst_excel_data)

def test_final_achieve():
    print('\n正在执行测试模块-----最终达成度')
    logger.debug('\n正在执行测试模块-----最终达成度')
    count_achieve, excel_achieve = test_get_data.get_final_achieve()
    excel_achieve = excel_achieve.astype(float)

    result = count_achieve == excel_achieve
    if False in result.values:
        logger.info('\n计算数据：\n{detail}', detail=count_achieve)
        logger.info('\nexcel数据：\n{detail}', detail=excel_achieve)
        for i in result.columns:
            detail = result[result[str(i)] == False].index.tolist()
            if len(detail) != 0:
                logger.info(str(i) + "达成度数据不一致的学号：{detail}", detail=detail)
    else:
        logger.debug("达成度测试通过")
    # pytest.assume(count_achieve.equals(excel_achieve))


def test_static():
    print('\n正在执行测试模块-----统计部分')
    logger.debug('\n正在执行测试模块-----统计部分')
    static_final_data, static_data_single_odd, static_data_all_odd, count_data, count_data_single_odd, count_data_all_odd = test_get_data.get_static_data()
    static_final_data = static_final_data.fillna(value=0)
    count_data = count_data.fillna(value=0)

    result = static_final_data == count_data
    if False in result.values:
        logger.info('\n计算数据：\n{detail}', detail=count_data)
        logger.info('\nexcel数据：\n{detail}', detail=static_final_data)
        for i in result.columns:
            detail = result[result[str(i)] == False].index.tolist()
            if len(detail) != 0:
                logger.info(str(i) + "统计数据不一致的项：{detail}", detail=detail)
    else:
        logger.debug("统计数据测试通过")

    result1 = static_data_single_odd == count_data_single_odd
    if False in result1.values:
        logger.info('\n计算数据：\n{detail}', detail=count_data_single_odd)
        logger.info('\nexcel数据：\n{detail}', detail=static_data_single_odd)
        for i in result1.columns:
            detail = result1[result1[str(i)] == False].index.tolist()
            if len(detail) != 0:
                logger.info(str(i) + "单项概率统计数据不一致的项：{detail}", detail=detail)
    else:
        logger.debug("单项概率统计数据测试通过")

    result2 = static_data_all_odd == count_data_all_odd
    if False in result2.values:
        logger.info('\n计算数据：\n{detail}', detail=count_data_all_odd)
        logger.info('\nexcel数据：\n{detail}', detail=static_data_all_odd)
        detail = result2[result2 == False].index.tolist()
        if len(detail) != 0:
            logger.info("统计数据不一致的项：{detail}", detail=detail)
    else:
        logger.debug("所有目标概率统计数据测试通过")
    # pytest.assume(static_final_data, count_data)
    # pytest.assume(static_data_single_odd, count_data_single_odd)
    # pytest.assume(static_data_all_odd, count_data_all_odd)


if __name__ == '__main__':
    pytest.main(['-vs'])