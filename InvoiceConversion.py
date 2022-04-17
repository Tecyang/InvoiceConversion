import sys
import os
import getopt
import logging
from InvoiceConversionOperation import InvoiceConversionOperation

import PdfToImg
import OperationWord


class InvoiceConversion:

    # 定义操作
    operation = None
    # 1、PDF地址
    pdfPath = './发票'
    # 2、需要储存图片的目录
    imagePath = './imgs'
    invoices = []
    wordName = 'test.docx'

    def addInvoiceImgToWord(self):
        doc = OperationWord.opWord(self.wordName)
        success = 0
        fail = 0
        print("开始处理要放到word的发票,共计{}个".format(len(self.invoices)))
        for invoice in self.invoices:
            print("正在查找发票图片{}".format(invoice))
            imgs = self.checkImgExist(invoice)
            print("查找发票图片结果{}".format(imgs))
            if len(imgs) > 0:
                OperationWord.addPicAndTitle(doc, invoice, imgs)
                success = success + 1
                print("发票{}成功".format(invoice))
            else:
                fail = fail+1
                print("发票{}失败".format(invoice))
        print("结束处理要放到word的发票,共计{}个,成功{}个，失败{}个".format(
            len(self.invoices), success, fail))

    def checkImgExist(self, invoiceNo):
        # imgs = os.path.join(self.imagePath,invoiceNo)
        imgs = []

        for dirpath, dirnames, filenames in os.walk(self.imagePath):
            for filename in filenames:
                print('判断文件是否为当前发票文件{},结果'.format(filename))
                if str(filename).startswith(invoiceNo):
                    print('True')
                    imgs.append(os.path.join(dirpath, filename))
        return imgs

    def init(self):
        print('进入init')
        # 获取构建参数
        opts, args = getopt.getopt(sys.argv[1:], "ho:")

        for op, value in opts:
            print("op is {} value is {} args is {}".format(op, value, args))
            if op == "-o":
                self.operation = value
                if InvoiceConversionOperation.IMG_FILL_WORD.value == self.operation:
                    self.invoices.append(args[0])
            elif op == "-h":
                self.usage()
                sys.exit()

    def usage(self):
        print("""
        发票转换脚本
        参数说明：
          -o:
            说明:执行的操作
          -h:
            说明:显示帮助
        """)


if __name__ == '__main__':

    conversion = InvoiceConversion()
    conversion.init()

    print(conversion.operation)
    # 发票转图片
    if InvoiceConversionOperation.PDF_TO_IMG.value == conversion.operation:
        PdfToImg.pyMuPDF_findPDF(conversion.pdfPath, conversion.imagePath)
    # 发票填充word
    elif InvoiceConversionOperation.IMG_FILL_WORD.value == conversion.operation:
        conversion.addInvoiceImgToWord()
