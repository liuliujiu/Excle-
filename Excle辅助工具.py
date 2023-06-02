import time

import xlwings as xw
import os

def space_check(wb,space=False):
    #定位到工作表
    sheet = wb.sheets[0]

    last_cell = sheet.used_range.last_cell
    nrows = sheet.used_range.last_cell.row
    ncols = sheet.used_range.last_cell.column

    print("工作表有：" + str(nrows) + "行" + "\n" + "工作表有：" + str(ncols) + "列")

    #枚举单元格数值并检查空格
    for i in range(1,nrows+1):
        for j in range(1,ncols+1):
            content = sheet.range(i,j).value
            if content == None:
                pass
            elif(" " in str(content)):
                print("单元格（{0}，{1}）存在空格".format(i,j))
                if space == True:
                    sheet.range(i, j).value = str(sheet.range(i, j).value).replace(" ","")
                    print("单元格（{0}，{1}）空格已经清除".format(i, j))
            else:
                pass


def line_check(wb,row,col,length):
    # 定位到工作表
    sheet = wb.sheets[0]

    last_cell = sheet.used_range.last_cell
    #获取有多少行
    nrows = sheet.used_range.last_cell.row

    for i in range(row,nrows +1):
        con_len = len(str(sheet.range(i,col).value))
        if con_len != length:
            print("单元格（{0}，{1}）长度不符合检查，实际长度为：{2}".format(i,col,con_len))


def table_merge(wb,dirname,row,col):
    excles = os.listdir(dirname)
    for i in excles:
        excle_name = dirname + "\\" + i

        local_wb = app.books.open(excle_name)
        time.sleep(3)
        local_wb.sheets[0].range(row,col).expand().copy(wb.sheets[0].range(41,1))






if __name__ == '__main__':
    # 设置excle打开方式等
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True

    # 打开excle
    wb = app.books.open('F:\\Test.xls')
    #space_check(wb, True)
    #line_check(wb, 3, 4, 10)

    table_merge(wb,"E:\\excle\\")
    # 保存
    #wb.sheets[1].range(3,1).expand().copy(wb.sheets[2].range(1,1))


    wb.save(path=None)
