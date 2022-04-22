import sys
import subprocess
import io
import os
import getopt
import json

from InvoiceConversionOperation import InvoiceConversionOperation
from CodeMessage import codeMessage
import PdfToImg
import OperationWord


class InvoiceConversion:

    # 定义操作
    operation = None
    # 1、PDF地址
    pdfPath = '../发票'
    # 2、需要储存图片的目录
    imagePath = '../imgs'
    wordPath = './'
    invoiceNo = ''
    invoices = []
    wordName = 'invoice.docx'

    def addInvoiceImgToWord(self):
        doc = OperationWord.opWord(self.wordPath, self.wordName)
        success = 0
        fail = 0
        codeMessage("开始处理要放到word的发票,共计{}个".format(len(self.invoices)))
        if len(self.invoices) > 0:
            for invoice in self.invoices:
                codeMessage("正在查找发票图片{}".format(invoice))
                imgs = self.checkImgExist(invoice)
                codeMessage("查找发票图片结果{}".format(imgs))
                if len(imgs) > 0:
                    OperationWord.addPicAndTitle(self.wordPath, self.wordName,
                                                 doc, invoice, imgs)
                    success = success + 1
                    codeMessage("发票{}成功".format(invoice))
                else:
                    fail = fail + 1
                    codeMessage("发票{}失败".format(invoice))
        codeMessage("结束处理要放到word的发票,共计{}个,成功{}个，失败{}个".format(
            len(self.invoices), success, fail))
        codeMessage("word文件路径为{}".format(
            os.path.abspath(os.path.join(self.wordPath, self.wordName))))

    def checkImgExist(self, invoiceNo):
        # imgs = os.path.join(self.imagePath,invoiceNo)
        imgs = []
        invoiceNo = str(invoiceNo).replace(' ', '')
        if invoiceNo == '':
            return []
        for dirpath, dirnames, filenames in os.walk(self.imagePath):
            for filename in filenames:
                # print('判断文件是否为当前发票文件{},结果'.format(filename))
                if str(filename).startswith(invoiceNo):
                    # print('True')
                    imgs.append(os.path.join(dirpath, filename))
        return imgs

    def init(self):
        # print('进入init')
        # 获取构建参数
        opts, args = getopt.getopt(sys.argv[1:], "ho:")

        for op, value in opts:
            # codeMessage("op is {} value is {} args is {}".format(op, value, args))
            # print(op + ' ' + value + ' {}'.format(args))
            if op == "-o":
                self.operation = value
                if len(args) > 0:
                    args0 = json.loads(str(args[0]))
                    self.pdfPath = args0['invoicePath']
                    self.imagePath = args0['imgPath']
                    self.wordPath = args0['wordPath']
                if len(args) > 1:
                    self.invoices.clear()
                    args1 = json.loads(str(args[1]))
                    self.invoiceNo = args1['invoiceNo']
                    if len(args1['invoiceNos']) > 0:
                        self.invoices = str(args1['invoiceNos']).split(',')
            elif op == "-h":
                self.usage()
                sys.exit()

    def usage(self):
        codeMessage("""
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
    # print(conversion.operation)
    # 发票转图片
    if InvoiceConversionOperation.PDF_TO_IMG.value == conversion.operation:
        PdfToImg.pyMuPDF_findPDF(conversion.pdfPath, conversion.imagePath)
    # 发票填充word
    elif InvoiceConversionOperation.IMG_FILL_WORD.value == conversion.operation:
        conversion.addInvoiceImgToWord()
    elif InvoiceConversionOperation.QUERY_INVOICE.value == conversion.operation:
        imgs = conversion.checkImgExist(conversion.invoiceNo)
        print(conversion.invoiceNo)
        print(imgs)
        if len(imgs) > 0:
            print('发票{}存在于图片资源目录中'.format(conversion.invoiceNo))
        else:
            print('发票{}不存在于图片资源目录中'.format(conversion.invoiceNo))
