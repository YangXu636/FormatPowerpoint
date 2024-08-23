from pptxOperationLibrary import pptxOp
import base64

if __name__ == "__main__":
    tmp = pptxOp(r"D:\Python\工具\FormatPowerpoint\sourceFile\第八章 循环系统(1).pptx")
    a = tmp.slidePictures(1)
    tmp.componentDelete(1, a[0])
