import shutil
import zipfile
import os
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
