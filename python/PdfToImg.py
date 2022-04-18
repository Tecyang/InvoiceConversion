import datetime
from email.mime import image
import os

import fitz  # fitz就是pip install PyMuPDF


def pyMuPDF_findPDF(pdfPath, imagePath):
    print("处理目录{}下的pdf,目标目录{}".format(pdfPath, imagePath))

    for file in os.listdir(pdfPath):
        file = os.path.join(pdfPath, file)
        print(file)
        # 判断当前目录是否为文件夹
        if os.path.isfile(file):
            # 进行转换
            pyMuPDF_fitz(pdfPath, file, imagePath)


def pyMuPDF_fitz(pdfPath, pdfFilePath, imagePath):
    startTime_pdf2img = datetime.datetime.now()  # 开始时间
    print("imagePath=" + imagePath)
    pdfDoc = fitz.open(pdfFilePath)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        zoom_x = 1.33333333  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 1.33333333
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):  # 判断存放图片的文件夹是否存在
            os.makedirs(imagePath)  # 若图片文件夹不存在就创建

            # 将图片写入指定的文件夹内
        if pg == 0:
            pix.writePNG(
                imagePath + '/' +
                str(pdfFilePath).replace(pdfPath, '').replace('pdf', 'png'))
        else:
            pix.writePNG(imagePath + '/' + str(pdfFilePath).replace(
                pdfPath, '').replace('.pdf', '_%s.png' % pg))

    endTime_pdf2img = datetime.datetime.now()  # 结束时间
    print('pdf2img时间=', (endTime_pdf2img - startTime_pdf2img).seconds)


if __name__ == "__main__":
    # 1、PDF地址
    pdfPath = './发票'
    # 2、需要储存图片的目录
    imagePath = './imgs'
    pyMuPDF_findPDF(pdfPath, imagePath)
