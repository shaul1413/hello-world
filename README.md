# hello-world
#随便写写，这是一个日历
#获取日历信息,并写入表格
import calendar
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

#创建工作薄
wb = openpyxl.Workbook()
#创建工作表，并标注清晰表名
for i in range(1,13):
    ws = wb.create_sheet(index=0,title=str(i)+'月')
    ws.cell(row=1,column=1).value = '2025年'+ str(i)+ '月'
    ws.cell(row=1,column=1).font = Font(name='Arial',size=20,color='FF0000')
#在表格中写入具体的时间，写入的信息就是日期
    for j in range(len(calendar.monthcalendar(2025,i))):
        for k in range(len(calendar.monthcalendar(2025,i)[j])):
            value = calendar.monthcalendar(2025,i)[j][k]

            if value==0:
                value=""
                ws.cell(row=j+9,column=k+1).value=value
            else:
                ws.cell(row=j+9,column=k+1).value=value
                ws.cell(row=j+9,column=k+1).font=Font(name='Arial',size=12)

    align = Alignment(horizontal='right', vertical='center')
    days = ['星期日','星期一','星期二','星期三','星期四','星期五','星期六']
    num = 0
    for l in range(len(days)):
        ws.cell(row=8,column=l+1).value=days[num]
        ws.cell(row=8,column=l+1).alignment=align
        ws.cell(row=8,column=l+1).font=Font(name='Arial',size=12)
        #将数字转化为字母 get_column_letter
        c_char = get_column_letter(l+1)
        #设置列高
        ws.column_dimensions[c_char].width = 10
        num+=1
    for l1 in range(8,14):
        #设置行高
        ws.row_dimensions[l1].height = 20

    fill = PatternFill(fill_type='solid',fgColor='00FF00')

    #对单元格底色进行着色
    for l2 in range(1,50):
        for l3 in range(1,50):
            ws.cell(row=l2,column=l3).fill = fill

    #添加图片
    img = Image(r"C:\Users\shulu.wang\Desktop\桌面\test\test.jpg")
    newsizes: tuple[int, int] = (400,400)
    img.width,img.height = newsizes
    #与顶部保持一些距离，好看一些
    ws.add_image(img,'I2')

wb.save(r"C:\Users\shulu.wang\Desktop\桌面\test\test11.xlsx")
