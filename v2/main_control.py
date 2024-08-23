import shutil
import zipfile
import os
import choose_ui
from pptxOperationLibrary import pptxOp
from betterSqlite3 import BetterSqlite3 as bs3
from betterTime import BetterTime as btime  # noqa: F401
from PIL import Image
import io


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
    root,
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
        root : tkinter的主窗口
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
    # 定义 模板、源文件、目标文件 的 pptxOp、bs3 对象
    mbPath = os.path.join(pathSrc, mbName + ".pptx")
    nfPath = os.path.join(pathSrc, nfName + ".pptx")
    targetPath = os.path.join(pathRlt, targetName + ".pptx")
    imagePath = os.path.join(pathSrc, "Images")
    mbPpt = pptxOp(mbPath)
    mbPptCount = mbPpt.slidesCount()
    mbPptCountLen = len(str(mbPptCount))
    if not os.path.exists(os.path.join(pathSrc, mbName + ".db")):
        mbDb = bs3(os.path.join(pathSrc, mbName + ".db"))
        # 模板文件分类
        mbDb.tableAdd("SlideType", {"Num": "INTEGER", "Type": "TEXT"})
        for i in range(1, mbPptCount + 1):
            # 分类整体页面
            slideBytes = mbPpt.slide2Bytes(i, "png")
            slidePng = Image.open(slideBytes)
            slidePng.show()
            slideType = choose_ui.getType(image=slideBytes, root=root)
            mbDb.dataInsert("SlideType", {"Num": i, "Type": slideType})
            mbDb.tableAdd(
                f"Slide{f'%0{mbPptCountLen}d' % i}",
                {"Texts": "TEXT", "Images": "BLOB", "Type": "TEXT"},
            )
            for j in mbPpt.slideTexts(i):
                # 分类单个文本框
                mbDb.dataInsert(
                    f"Slide{f'%0{mbPptCountLen}d' % i}",
                    {
                        "Texts": j,
                        "Images": None,
                        "Type": choose_ui.getType(text=j, root=root),
                    },
                )
            for j in mbPpt.slidePictures(i):
                # 分类单个图片
                mbDb.dataInsert(
                    f"Slide{f'%0{mbPptCountLen}d' % i}",
                    {
                        "Texts": None,
                        "Images": j,
                        "Type": choose_ui.getType(image=io.BytesIO(j), root=root),
                    },
                )
            slidePng.close()
        mbDb.dbClose()
        del mbDb
    mbDb = bs3(os.path.join(pathSrc, mbName + ".db"))
    nfPpt = pptxOp(nfPath)
    targetPpt = pptxOp(targetPath)
    nfDb = bs3(os.path.join(pathSrc, targetName + ".db"))
    # 分类源文件内容
    nfDb.tableAdd("SlideType", {"Num": "INTEGER", "Type": "TEXT"})
    nfPptCount = nfPpt.slidesCount()
    nfPptCountLen = len(str(nfPptCount))
    targetPpt.fileNew()
    for i in range(max(1, startIndex), min(nfPptCount + 1, endIndex + 1)):
        try:
            logConsole(f"正在处理第{i}页...")
            print(f"正在处理第{i}页...")
            nfDb.tableAdd(
                f"Slide{f'%0{nfPptCountLen}d' % i}",
                {"Texts": "TEXT", "Images": "BLOB", "Type": "TEXT"},
            )
            # 获取所有文本和图片
            slideTexts = nfPpt.slideTexts(i)
            slideImages = nfPpt.slidePictures(i)
            print(f"{len(slideTexts) = } {slideTexts = } {len(slideImages) = }")
            nfPpt.slide2Image(
                i, imagePath + f"\\{nfName}", f"Slide{f'%0{nfPptCountLen}d' % i}", "png"
            )
            # 显示整体图片
            slidePng = Image.open(
                os.path.join(
                    imagePath + f"\\{nfName}", f"Slide{f'%0{nfPptCountLen}d' % i}.png"
                )
            )
            slidePng.show()
            # 分类整体内容
            slideType = choose_ui.getType(image=nfPpt.slide2Bytes(i, "png"), root=root)
            nfDb.dataInsert("SlideType", {"Num": i, "Type": slideType})
            if slideType == "无用":
                continue
            mbNum = mbDb.dataSelect("SlideType", f'Type == "{slideType}"', ["Num"])[0][
                0
            ]  # 获取模板页码
            targetPpt.slideCopy(
                mbPpt.file_path, mbNum, targetPpt.slidesCount() + 1
            )  # 复制模板页
            for j in slideTexts:
                Type = choose_ui.getType(text=j, root=root)
                nfDb.dataInsert(
                    f"Slide{f'%0{nfPptCountLen}d' % i}",
                    {"Texts": j, "Images": None, "Type": Type},
                )
            for j in set(
                [
                    k[0]
                    for k in nfDb.dataSelect(
                        f"Slide{f'%0{nfPptCountLen}d' % i}", "Images IS NULL", ["Type"]
                    )
                ]
            ):  # 遍历所有类型
                if j == "无用":
                    continue
                texts = [
                    k[0]
                    for k in nfDb.dataSelect(
                        f"Slide{f'%0{nfPptCountLen}d' % i}", f'Type == "{j}"', ["Texts"]
                    )
                ]  # 获取所有文本
                print(f"Type = {j}    {texts = }")
                try:
                    targetPpt.textChange(
                        targetPpt.slidesCount(),
                        mbDb.dataSelect(
                            f"Slide{f'%0{mbPptCountLen}d' % mbNum}",
                            f'Type == "{j}"',
                            ["Texts"],
                        )[0][0],
                        "".join(texts),
                    )  # 改变文本
                except Exception as e:
                    logConsole(f"第{i}页{j}替换失败！    {e = }", lvl="Warning")
                    print(f"第{i}页{j}替换失败！    {e}")
            targetPpt.componentDelete(
                i, list(set(targetPpt.slideTexts(i)) & set(mbPpt.slideTexts(mbNum)))
            )  # 删除模板页中未改变的多余的文本框
            for j in slideImages:
                Type = choose_ui.getType(image=io.BytesIO(j), root=root)
                nfDb.dataInsert(
                    f"Slide{f'%0{nfPptCountLen}d' % i}",
                    {"Texts": None, "Images": j, "Type": Type},
                )
            for j in set(
                [
                    k[0]
                    for k in nfDb.dataSelect(
                        f"Slide{f'%0{nfPptCountLen}d' % i}", "Texts IS NULL", ["Type"]
                    )
                ]
            ):  # 遍历所有类型
                if j == "无用":
                    continue
                images = [
                    k[0]
                    for k in nfDb.dataSelect(
                        f"Slide{f'%0{nfPptCountLen}d' % i}",
                        f'Type == "{j}"',
                        ["Images"],
                    )
                ][0]  # 仅取第一个图片
                try:
                    targetPpt.pictureChange(
                        targetPpt.slidesCount(),
                        mbDb.dataSelect(
                            f"Slide{f'%0{mbPptCountLen}d' % mbNum}",
                            f'Type == "{j}"',
                            ["Images"],
                        )[0][0],
                        io.BytesIO(images),
                    )
                except Exception as e:
                    logConsole(f"第{i}页{j}替换失败！    {e = }", lvl="Warning")
                    print(f"第{i}页{j}替换失败！    {e}")
                print(
                    list(
                        set(targetPpt.slidePictures(i))
                        & set(mbPpt.slidePictures(mbNum))
                    )
                )
            targetPpt.componentDelete(
                i,
                list(set(targetPpt.slidePictures(i)) & set(mbPpt.slidePictures(mbNum))),
            )  # 删除模板页中未改变的多余的图片
            del slidePng
            logSucceed(f"{i = }")
        except Exception as e:
            logFailed(f"{i = }")
            logConsole(f"格式化第{i}页失败！    {e = }", lvl="Error")
            print(f"格式化第{i}页失败！    {e}")
            continue
    nfDb.dbClose()
    del nfDb
    del nfPpt
    del targetPpt
    del mbPpt
    del mbDb
    logConsole(f"格式化PowerPoint文件成功！\n    {pathRlt = }\n    {targetName = }")
    return
