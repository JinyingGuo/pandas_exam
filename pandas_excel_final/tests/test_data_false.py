import pytest
from log.log import logger
from src.main import object_all_out

@pytest.fixture(autouse=True)
def set_up():
    yield
    print("\n测试结束")

# 作业页帧总分验证
def test_out():
    print('\n正在执行测试模块-----作业页帧总分')
    logger.debug('\n正在执行测试模块-----作业页帧总分')
    data_out = object_all_out[0]
    count_result1 = data_out.loc[20162573, :].sum()
    count_result2 = data_out.loc[20172725, :].sum()

    pytest.assume(count_result1 == 20)
    pytest.assume(count_result2 == 28)

# 作业页帧达成度验证
def test_achieve():
    print('\n正在执行测试模块-----作业页帧达成度')
    logger.debug('\n正在执行测试模块-----作业页帧达成度')


if __name__ == '__main__':
    pytest.main(['-vs'])