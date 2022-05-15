from src import function
import pandas as pd

def count_object_target(object_all_scores, object_all_param, df_weight):
    object_evaluation_out = list()              # 保存得分
    object_evaluation_achieve = list()          # 保存达成度
    object_param = list()                       # 计算完总分的参数表
    for i in range(len(object_all_scores)):
        single_achieve, single_out, count_param = function.count_object_evaluation(object_all_param[i], object_all_scores[i], df_weight,
                                                         df_weight.index.values)
        object_evaluation_achieve.append(single_achieve)
        object_evaluation_out.append(single_out)
        object_param.append(count_param)
    return object_evaluation_achieve, object_evaluation_out, object_param


def concat_data(df_weight, object_evaluation_achieve, unobject_scores, student_id):
    all_achieve_data = pd.DataFrame(data=None, index=student_id)
    for i in df_weight.index:
        all_target_data = function.count_object_evaluate(object_evaluation_achieve, unobject_scores, i, all_achieve_data, df_weight)
    return all_target_data

def count_final_achieve(all_target_data, df_weight, student_id):
    # 计算达成度
    target_achieve = pd.DataFrame(index=student_id)
    for i in range(len(all_target_data)):
        achievement = function.count_evaluate(all_target_data[i], df_weight, df_weight.index.values[i])
        target_achieve.insert(len(target_achieve.columns), '课程目标' + str(i + 1), achievement)
    # 计算均值
    target_achieve.loc['平均值'] = target_achieve.apply(lambda x: round(x.mean(), 3), axis=0)
    return target_achieve

def count_class_avg_ach(object_evaluation_out, object_all_param, unobject_scores, unobject_param, df_weight, class_id, everyclass_student):
    class_ach = list()                      # 存放班级达成度
    class_avg = list()                      # 存放班级平均值
    # 客观评价部分
    class_objachieve = pd.Series(data=None, index=df_weight.index, dtype=object)
    class_objaverage = pd.Series(data=None, index=df_weight.index, dtype=object)
    for i in range(len(object_evaluation_out)):
        average, achieve = function.count_class_object(object_evaluation_out[i], object_all_param[i], class_id, everyclass_student, class_objachieve, class_objaverage)
        class_ach.append(achieve)
        class_avg.append(average)
    class_unobjachieve = pd.Series(data=None, index=df_weight.index, dtype=object)
    class_unobjaverage = pd.Series(data=None, index=df_weight.index, dtype=object)
    # 主观部分
    for i in range(len(unobject_param)):
        class_achieve, class_average = function.count_class_unobject(unobject_scores[i], unobject_param[i], everyclass_student, class_unobjachieve, class_unobjaverage)
        class_achieve = class_achieve.dropna(axis=1, how='all')
        class_achieve.columns = class_id
        class_ach.append(class_achieve.T)
        class_average = class_average.dropna(axis=1, how='all')
        class_average.columns = class_id
        class_avg.append(class_average.T)
    return class_ach, class_avg

all_everyclass_achieve = list()
def concat_everyclass_achieve(class_ach, class_id, df_weight):
    for i in class_id:
        class_achieve = pd.DataFrame(data=None, index=df_weight.index)
        class_achieve = function.concat_class_achieve(class_ach, class_achieve, i, df_weight)
        all_everyclass_achieve.append(class_achieve)
    return all_everyclass_achieve