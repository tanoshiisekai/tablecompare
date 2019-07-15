from pyexcel_xls import get_data as get_data_xls
from pyexcel_xlsx import get_data as get_data_xlsx

dirname = "C:/datacompare/"  # 目标文件的所在目录

file1 = "file1.xlsx"      # 文件1的名字（要参照的表格）
tablename1 = "Sheet1"         # 文件1中要参照的表名
col1name = ['B', 'E', 'F']  # 文件1中的有效列（注意排序对结果有影响）
row1 = 2                    # 文件1从第几行数据开始读取，从0开始计数
rowcount1 = 7

file2 = "file2.xls"    # 文件2的名字（要核对的表格）
tablename2 = "Sheet2"          # 文件2中要核对的表名
col2name = ['A', 'E', 'B']        # 文件2中的有效列（注意排序对结果有影响）
row2 = 3                             # 文件2从第几行数据开始读取，从0开始计数
rowcount2 = 41

mainkeycol1 = 0             # 提取结果中的备选关键字列1
mainkeycol2 = 1             # 提取结果中的备选关键字列2


def getcolnum(colname):
    resultlist = []
    for col in colname:
        thesum = 0
        length = len(col)
        loop = length - 1
        while loop >= 0:
            thesum = thesum + (ord(col[length-loop-1])-ord('A') + 1) * (26 ** loop)
            loop = loop - 1
        resultlist.append(thesum - 1)
    return resultlist


def listequal(list1, list2, keycol1, keycol2):
    if len(list1) != len(list2):
        return None
    else:
        for x in range(0, len(list1)):
            if list1[x] != list2[x]:
                if list1[keycol1] == list2[keycol1] or list1[keycol2] == list2[keycol2]:
                    return tuple(((list1[keycol1], list1[keycol2], list1[x]),
                                 (list2[keycol1], list2[keycol2], list2[x])))
                return tuple((False, list1[x], list2[x]))
        return True

col1 = getcolnum(col1name)
col2 = getcolnum(col2name)

if file1.endswith(".xls"):
    f1 = get_data_xls(dirname + file1)[tablename1][row1:]
elif file1.endswith(".xlsx"):
    f1 = get_data_xlsx(dirname + file1)[tablename1][row1:]
table1list = []
for i in range(0, rowcount1):
    table1list.append([][:])
for c in col1:
    for d in range(0, rowcount1):
        table1list[d].append(f1[d][c])

if file2.endswith(".xls"):
    f2 = get_data_xls(dirname + file2)[tablename2][row2:]
elif file2.endswith(".xlsx"):
    f2 = get_data_xlsx(dirname + file2)[tablename2][row2:]
table2list = []
for i in range(0, rowcount2):
    table2list.append([][:])
for c in col2:
    for d in range(0, rowcount2):
        table2list[d].append(f2[d][c])

print(table1list)
print(table2list)

reflist = table1list
aimlist = table2list

passcount = 0
for aim in aimlist:
    iflag = 0
    for ref in reflist:
        flag = listequal(aim, ref, mainkeycol1, mainkeycol2)
        if flag is True:
            iflag = 1
            passcount = passcount + 1
            break
        if iflag == 0:
            if flag[0] is not False:
                print("E: ", "\n", aim, "\n", ref, "\n", "===>", flag)
if passcount == len(aimlist):
    print("All Passed.")
print("Pass Count:", passcount)


