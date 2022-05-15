import os

# def get_excel_name():
#     filelist = []
#     for root, dirs, files in os.walk('..\data'):
#         for file in files:
#             if file.endswith(".xlsx") or file.endswith(".xls"):
#                 file_name = os.path.join(root, file)
#                 filelist.append(file_name)
#     return filelist

def get_excel_name():
    filelist = []
    for root, dirs, files in os.walk('..\data'):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                filelist.append(file)
        file_key = list(range(1, len(filelist) + 1))
        path = dict(zip(file_key, filelist))
        path[len(path) + 1] = "处理全部excel"
        for key, value in path.items():
            print(key, value)
        print("请输入想要处理的excel序号:")
        a = int(input())
        filelist.clear()
        if a == len(path):
            for file in files:
                file_name = os.path.join(root, file)
                filelist.append(file_name)
        else:
            filelist.append(os.path.join(root, path[a]))
    return filelist
    # return filelist


file_name = get_excel_name()
print(file_name)
# l1 = list(range(1, len(filelist)+1))
# path = dict(zip(l1, filelist))
# for key, value in path.items():
#     print(key, value)
# print("请输入想要处理的excel序号:")
# a = int(input())
# print(path[a])