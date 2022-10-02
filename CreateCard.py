# CardCreator 
# Generate stored value card image according to table data
# 储值卡生成器 - 根据表格数据生成储值卡图片
# hegeo@foxmail.com
# 2022-10
# install libs:
# pip install Pillow
# pip install pystrich
# pip install xlrd



from PIL import Image, ImageDraw, ImageFont
from pystrich.code128 import Code128Encoder
import time        
import xlrd
import time
import math
import re
import os,sys
os.chdir(sys.path[0])



# 删除文件操作
def del_files0(dir_path):
    shutil.rmtree(dir_path)

#删除单文件
def del_files(dir_path):
    if os.path.isfile(dir_path):
        try:
            os.remove(dir_path) 
        except BaseException as e:
            print(e)
    elif os.path.isdir(dir_path):
        file_lis = os.listdir(dir_path)
        for file_name in file_lis:
            # if file_name != 'wibot.log':
            tf = os.path.join(dir_path, file_name)
            del_files(tf)
                        
#数据提取处理
data = xlrd.open_workbook("./card.xls") # 只支持xls文件
table = data.sheets()[0] # 打开第一张表
nrows = table.nrows # 获取表的行数

#制卡方法
def makecard(cid,pin,cash):
    
    #数据格式处理
    cash = cash
    cash00 = '$'+str(math.floor(int(cash))) +'.00'
    mtext = 'Code Card  ...'+str(cid)[12:16]
    #CardID空格4间隔
    cid = cid
    text_list = re.findall(".{4}",cid)
    cardid = " ".join(text_list)

    #PIN空格间隔
    pin = ' '.join(str(math.floor(int(pin))))
    barcodetext = cid

    #cash = '$250.00'
    #mtext = 'Code Card  ...0001'
    #cardid = '6543 1234 6789 1234'
    #pin = '6 6 0 0 0 1 0 1'
    #barcodetext = '6193766785648006'

    #图片文件路径    
    img = "./resource/muban.png"  # 图片模板
    new_img = "./card/tempcard.png"  # 生成的图片
    compress_img = "./card/"+cid+".png"  # 压缩后的图片
    #compress_img = "./card/newcard.png"  # 压缩后的图片

    #字体风格
    cash_type = ImageFont.truetype("./resource/SourceHanSans-Bold.otf",160)
    mtext_type = ImageFont.truetype("./resource/SourceHanSans-Bold.otf",55)
    mcash_type = ImageFont.truetype("./resource/SourceHanSans-Medium.otf",45)
    btext_type = ImageFont.truetype("./resource/SourceHanSans-Medium.otf",65)
    barcode_type = ImageFont.truetype("./resource/code128.ttf",170)
    blackcolor = "#000000"
    bluecolor = "#317699"
    graycolor = "#9299ac"
    redcolor = "#e02f49"

    #条形码
    encoder = Code128Encoder(barcodetext)
    encoder.save( "./bartest.png" )

    #打开模板
    image = Image.open(img)
    draw = ImageDraw.Draw(image)
    width, height = image.size

    #大金额
    if cash<99:
        cash_x = 270
        cash_y = 1050
    elif cash<999:
        cash_x = 230
        cash_y = 1050
    else:
        cash_x = 190
        cash_y = 1050    
    draw.text((cash_x, cash_y), cash00, bluecolor, cash_type)

    #小金额
    if cash<99:
        cash2_x = 880
        cash2_y = 1390
    elif cash<999:
        cash2_x = 860
        cash2_y = 1390
    else:
        cash2_x = 840
        cash2_y = 1390
    draw.text((cash2_x, cash2_y), cash00, graycolor, mcash_type)

    #标题文字
    mtext_x = 310
    mtext_y = 930
    draw.text((mtext_x, mtext_y), mtext, blackcolor, mtext_type)

    #卡号
    cardid_x = 50
    cardid_y = 1550
    draw.text((cardid_x, cardid_y), cardid, blackcolor, btext_type)

    #PIN码
    pin_x = 60
    pin_y = 1730
    draw.text((pin_x, pin_y), pin, blackcolor, btext_type)

    #条形码处理
    #barcode_x = 90
    #barcode_y = 1860
    #draw.text((barcode_x, barcode_y), barcodetext, blackcolor, barcode_type)
    bartempimg = Image.open("./bartest.png") 
    region = bartempimg.crop((10,14,420,112))  # 0,0,50,50左上右下裁剪坐标
    region = region.resize((920,150), Image.ANTIALIAS)
    region.save("./barcode.png")
    del_files("./bartest.png")
    
    #保存图片
    image.save(new_img, 'png')
    bigimg = Image.open("./card/tempcard.png") 
    barcodeimg = Image.open("./barcode.png") 
    bigimg.paste(barcodeimg,(85,1880))
    bigimg.save(new_img, 'png')
    del_files("./barcode.png")

    #压缩图片
    sImg = Image.open(new_img)
    w, h = sImg.size
    width = int(w / 2)
    height = int(h / 2)
    dImg = sImg.resize((width, height), Image.ANTIALIAS)
    dImg.save(compress_img)
    del_files("./card/tempcard.png")

#流水线
for i in range(nrows):
   if i == 0: # 跳过第一行
       continue
   #print(table.row_values(i)[:3]) # 取前4列数
   makecard(table.row_values(i)[0],table.row_values(i)[1],table.row_values(i)[2])

