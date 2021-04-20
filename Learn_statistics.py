# -*- coding: utf-8 -*-

title = """     ———————————————————————  青年大学习统计系统  ———————————————————————
输入大学习系统导出文件，输出两个文件：
1. 文件名：m季n期大学习参与率排名（所属组织，团员数，第n期大学习参与人数，第n期大学习参与率）。 各年级前5名及学习率超过100%的班级'整行字体'标红，后5名标蓝；
2. 文件名：m季n期学习情况（班级，姓名，学习情况）。 未学习的学生'整行底色'标黄。

author：任鑫英
mail：flerken@stu.ncst.edu.cn
github：https://github.com/XingKong0927
"""

# excel文件上色
def color_execl(data0, save_file_name):
    import pandas as pd
    count_row = len(data0)
    writer = pd.ExcelWriter("result\\{}".format(save_file_name), engine='xlsxwriter')
    data0.to_excel(writer, sheet_name = 'Sheet1', index = False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'bg_color':   '#ffff00',
                                'font_color': '#9C0006'})
    worksheet.conditional_format('A1:C{}'.format(count_row+1), {'type':     'cell',
                                                            'criteria': '==',
                                                            'value':    '"未学习"',
                                                            'format':   format1})
    writer.save()

if __name__=='__main__':
    
    print(title)
    print("———————————————————————  程序开始  ———————————————————————")
    print("重要提示：请将导出文件放入本程序解压目录 data文件夹 中 ！！！")

    # import time
    import datetime
    import os
    import pandas as pd


    file_name = input("\n请输入大学习系统导出文件名称（eg：161830245419导出数据）：")
    print("\n⇩ ⇩  接下来请输入本次大学习 季数 和 期数  ⇩ ⇩")
    time_season = input("请输入本次学习季数（eg：3）：")
    time_stage = input("请输入本次学习期数（eg：4）：")
    # file_name = '161830245419导出数据'

    starttime = datetime.datetime.now()
    print("\n程序开始运行，请稍候...\n")

    data0 = pd.read_csv("data\\{}.csv".format(file_name), engine='python')
    data1 = pd.DataFrame(data = data0)
    data1.columns = ['name', 'id_card', 'time', 'class', 'type', 'Municipal', 'county', 'town']      # 更换表头

    """文件1：m季n期大学习参与率排名"""
    data10 = data1["class"]
    data11 = pd.read_csv("source\\第n期大学习参与率排名.csv", engine = 'python')
    data12 = pd.DataFrame(data = data11)
    
    # 进度条代码(python可用，输出到命令行会出问题)
    # len_bar = 100           # 进度条长度
    # len_data10 = len(data10)
    # beishu = int(len_data10/len_bar)
    # start = time.perf_counter()
    for row0 in range(len(data10)):
        # # 进度条代码
        # a='*'*int(row0/beishu)
        # b='.'*int((len_data10-row0)/beishu)
        # c=(row0/len_data10)*100
        # dur=time.perf_counter()-start
        # print("\r正在处理参与率排名：{:^3.0f}%[{}->{}]{:.2f}s".format(c,a,b,dur), end="")
        
        branch0 = data10[row0]
        for row1 in range(len(data12)):
            if branch0 == data12.loc[row1]['所属组织']:
                data12.loc[row1, '第n期大学习参与人数'] += 1
                break
    for row1 in range(len(data12)):
        data12.loc[row1, '第n期大学习参与率'] = '{:.2%}'.format(data12.loc[row1]['第n期大学习参与人数']/data12.loc[row1]['团员数'])
    data12.columns = ['所属组织', '团员数', '第{}期大学习参与人数'.format(time_stage), '第{}期大学习参与率'.format(time_stage)]      # 更换表头
    data12.to_excel("result\\{}季{}期大学习参与率排名.xlsx".format(time_season, time_stage), index = None)


    """文件2：m季n期学习情况"""
    save_file_name = '{}季{}期学习情况.xlsx'.format(time_season, time_stage)

    data21 = pd.read_csv("source\\成员团关系.csv", engine = 'python')
    data22 = pd.DataFrame(data = data21)

    len_data22 = len(data22)
    for row in range(len_data22):

        id_card = data22.loc[row]['证件号码']
        name = data22.loc[row]['姓名']
        sign = 0
        for row0 in range(len(data1)):
            if id_card == data1.loc[row0]['id_card'] and name == data1.loc[row0]['name']:
                sign = 1
                break
        if sign:
            data22.loc[row, '学习情况'] = '已学习'
        else:
            data22.loc[row, '学习情况'] = '未学习'
    data23 = data22.drop(['证件号码'], axis=1)
    # data23.to_excel("m季n期学习情况.xlsx", index = None)

    # 上色并保存文件
    color_execl(data23, save_file_name)

    endtime = datetime.datetime.now()
    print("\n————————— 程序结束，运行时间：{}。已输出文件于 result文件夹，按任意键退出 —————————".format(str(endtime - starttime)[:-7]))
    os.system('pause')



