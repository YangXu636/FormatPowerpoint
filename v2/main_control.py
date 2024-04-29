import shutil
import zipfile
import os
import choose_ui  # noqa: F401
from pptxOperationLibrary import pptxOp
from betterSqlite3 import BetterSqlite3 as bs3
from betterTime import BetterTime as btime  # noqa: F401
# import win32com.client
# import pywintypes
# import time
# from pptx import Presentation
# from pptx.dml.color import _NoneColor
# from pptx.enum.shapes import MSO_SHAPE_TYPE
# from pptx.util import Emu, Pt
# import re
# from PIL import Image


def copy_file(src, dst):
    shutil.copy(src, dst)


def remove_file(path):
    if os.path.isdir(path):
        shutil.rmtree(path)
        os.mkdir(path)
    else:
        os.remove(path)


def ZipExtract(zip_file, dir):
    if not os.path.exists(dir):
        os.mkdir(dir)
    with zipfile.ZipFile(zip_file, "r") as zip_ref:
        zip_ref.extractall(path=dir)


def ZipCompress(zip_file, dir):
    if os.path.exists(zip_file):
        os.remove(zip_file)
    zip = zipfile.ZipFile(zip_file, "w", zipfile.ZIP_DEFLATED)
    for path, dirnames, filenames in os.walk(dir):  # type:ignore
        # 去掉目标跟路径，只对目标文件夹下边的文件及文件夹进行压缩
        fpath = path.replace(dir, "")
        for filename in filenames:
            zip.write(os.path.join(path, filename), os.path.join(fpath, filename))
    zip.close()


def is_number(s) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        pass
    try:
        import unicodedata

        for i in s:
            unicodedata.numeric(i)
            return True
    except (TypeError, ValueError):
        pass
    return False


def format_powerpoint(
    pathSrc: str,
    pathRlt: str,
    mbName: str,
    nfName: str,
    targetName: str,
    startIndex: int,
    endIndex: int,
    logConsole,
    logSucceed,
    logFailed,
):
    """
    格式化PowerPoint文件

    param:
        pathSrc: 源文件路径
        pathRlt: 结果文件路径
        mbName: 模板文件名
        nfName: 源文件名
        targetName: 目标文件名
        startIndex: 开始页码
        endIndex: 结束页码
        logConsole: 日志输出函数
        logSucceed: 成功日志输出函数
        logFailed: 失败日志输出函数

    return:
        None
    """
    logConsole(
        f"开始格式化PowerPoint文件...\n    {pathSrc = }\n    {pathRlt = }\n    {mbName = }\n    {nfName = }\n    {targetName = }\n    {startIndex = }\n    {endIndex = }"
    )
    logConsole("开始分类模板文件内容...")
    mbPath = os.path.join(pathSrc, mbName + ".pptx")
    nfPath = os.path.join(pathSrc, nfName + ".pptx")
    targetPath = os.path.join(pathRlt, targetName + ".pptx")
    mbPpt = pptxOp(mbPath)  # noqa: F841
    nfPpt = pptxOp(nfPath)
    targetPpt = pptxOp(targetPath)
    nfDb = bs3(os.path.join(pathSrc, targetName + ".db"))
    nfDb.tableAdd("SlideType", {"Num": "INTEGER", "Type": "TEXT"})
    nfPptCount = nfPpt.slidesCount()
    nfPptCountLen = len(str(nfPptCount))
    for i in range(max(1, startIndex), min(nfPptCount + 1, endIndex)):
        if i < startIndex or i > endIndex:
            continue
        logConsole(f"正在处理第{i}页...")
        print(f"正在处理第{i}页...")
        nfDb.tableAdd(
            f"Slide{f'%0{nfPptCountLen}d' % i}",
            {"Texts": "TEXT", "Images": "BLOB", "Type": "TEXT"},
        )
        slideTexts = nfPpt.slideTexts(i)
        slideImages = nfPpt.slidePictures(i)
        print(f"{len(slideTexts) = } {slideTexts = } {len(slideImages) = }")
        for j in slideTexts:
            nfDb.dataInsert(
                f"Slide{f'%0{nfPptCountLen}d' % i}",
                {"Texts": j, "Images": None, "Type": None},
            )
        for j in slideImages:
            nfDb.dataInsert(
                f"Slide{f'%0{nfPptCountLen}d' % i}",
                {"Texts": None, "Images": j, "Type": None},
            )
    nfDb.dbClose()
    del nfDb
    targetPpt.newFile()
    return
