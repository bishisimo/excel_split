import os

import xlrd
import xlwt
from xlwt import Worksheet


def excel_read(path: str, title_num: int = 4):
    """

    :param path: 文件完整路径
    :param title_num: 标题数量
    :return: 包含标题与数据的列表
    """
    if not os.path.exists(path):
        return False
    title_rows = []
    data = []
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    row_all_num = sheet.nrows
    for i in range(row_all_num):
        if i < title_num:
            # print('标题:'+str(i), sheet.row_values(i))
            row: list = sheet.row_values(i)
            row.pop(2)
            title_rows.append(row)
        else:
            # print('内容:', sheet.row_values(i))
            row: list = sheet.row_values(i)
            row.pop(2)
            data.append(row)
    return [title_rows, data]


def set_style(font_name: str = 'Times New Roman', height: int = 220, alignment: bool = True):
    """

    :param font_name: 字体名
    :param height: 字体高度
    :param alignment: 是否居中
    :return: Excel样式
    """
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = font_name  # 'Times New Roman'
    font.color_index = 4
    font.height = height
    style.font = font

    if alignment:
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al

    return style


def excel_save(root_path: str, title_rows: list, data_rows: list):
    """

    :param root_path: 储存生成文件的根路径
    :param title_rows: 包含标题的行
    :param data_rows: 包含数据的行
    :return:
    """
    if not os.path.exists(root_path):
        os.mkdir(root_path)

    title_num = len(title_rows)
    for index, person in enumerate(data_rows):
        person_name = person[0]
        file_path = os.path.join(root_path, person_name + '.xls')
        print('---->生成表格开始:',person[0])
        wb = xlwt.Workbook()
        sheet: Worksheet = wb.add_sheet('月度汇总', cell_overwrite_ok=True)
        print('-------->加入标题开始')
        # 加入标题
        for i, row in enumerate(title_rows):
            for j, item in enumerate(row):
                if i == 1:
                    sheet.col(j).width = 256 * 12
                if i == 2 and (j < 8 or j > 10):
                    sheet.write_merge(i, i + 1, j, j, item, set_style())
                elif i == 2 and j == 9:
                    sheet.write_merge(i, i, j, j + 1, item, set_style())
                elif i == 0 and j == 0:
                    sheet.write_merge(i, i, 0, len(row) - 1, item, set_style(height=440, alignment=False))
                elif i == 1 and j == 0:
                    sheet.write_merge(i, i, 0, len(row) - 1, item, set_style(alignment=False))
                elif i > 2:
                    sheet.write(i, j, item, set_style())
        print('-------->加入标题完成')
        # 加入数据
        print('-------->加入数据开始')
        for j, it in enumerate(person):
            sheet.write(title_num, j, it, set_style())
        print('-------->加入数据完成')
        print('---->生成表格完成')
        print(file_path)
        with open(file_path,'wb')as file:
            print('---->写入表格开始')
            wb.save(file)
        print('---->写入表格完成')


if __name__ == '__main__':
    for name in os.listdir('./'):
        if not ('.xlsx' in name or 'xls' in name):
            continue
        if not os.path.exists('out'):
            os.mkdir('out')
        file_path = os.path.join('./out')
        try:
            context = excel_read(name)
            print('解析完毕,正在生成文件...')
            excel_save(file_path, *context)
            print('生成完毕')
        except Exception as e:
            print('解析过程产生异常:', e)
            input()
        break


    print('\n\n文件解析完毕,请查收')
    input()
