import pptxOperationError as pptxOpError
import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa
from pptx.util import Emu
import win32com.client
import os
from typing import overload


class pptxOp:
    """Better Powerpoint Operation Library"""

    def __init__(self, file_path, t=10):
        self.file_path = file_path
        self.t = t

    @overload
    def newFile(self) -> None:
        """
        new_file()

        Create a new Powerpoint file.
        """
        t = self.t
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")

                prs = powerpoint.Presentations.Add()
                prs.SaveAs(self.file_path)
                prs.Close()
                powerpoint.Quit()
                return
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeCreatedError(self.file_path)

    @staticmethod
    def newFile(file_path: str, t: int = 10) -> None:
        """
        new_file(file_path [,t=10])

        Create a new Powerpoint file.
        """

        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")

                prs = powerpoint.Presentations.Add()
                prs.SaveAs(file_path)
                prs.Close()
                powerpoint.Quit()
                return
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeCreatedError(file_path)

    @overload
    def pptFileConversion(
        self, new_file_format: str, new_file_path: str = None
    ) -> None:
        """
        ppt_file_conversion(new_file_format,  [,new_file_path=None])

        Convert Powerpoint file to new_file_format.

        new_file_format: 'ppt', 'pptx', 'pdf', 'png', 'jpg', 'gif', 'tiff', 'bmp'
        new_file_path: new file path.
        """
        t = self.t
        if not os.path.exists(self.file_path):
            raise pptxOpError.FileNotFoundError(self.file_path)
        if new_file_path is None:
            new_file_path = os.path.splitext(self.file_path)[0] + "." + new_file_format
        else:
            new_file_path = (
                new_file_path
                + os.path.basename(self.file_path).split(".")[0]
                + "."
                + new_file_format
            )
        if not os.path.exists(os.path.dirname(new_file_path)):
            os.makedirs(os.path.dirname(new_file_path))
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                prs.SaveAs(new_file_path, FileFormat=new_file_format)
                prs.Close()
                powerpoint.Quit()
                return
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeConvertedError(self.file_path, new_file_format)

    @staticmethod
    def pptFileConversion(
        file: str, new_file_format: str, new_file_path: str = None, t: int = 10
    ) -> None:
        """
        ppt_file_conversion(file, new_file_format,  [,new_file_path=None] [,t=10])

        Convert Powerpoint file to new_file_format.

        new_file_format: 'ppt', 'pptx', 'pdf', 'png', 'jpg', 'gif', 'tiff', 'bmp'
        new_file_path: new file path.
        t is the number of repeated trial and error attempts.
        """
        if not os.path.exists(file):
            raise pptxOpError.FileNotFoundError(file)
        if new_file_path is None:
            new_file_path = os.path.splitext(file)[0] + "." + new_file_format
        else:
            new_file_path = (
                new_file_path
                + os.path.basename(file).split(".")[0]
                + "."
                + new_file_format
            )
        if not os.path.exists(os.path.dirname(new_file_path)):
            os.makedirs(os.path.dirname(new_file_path))
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(file, WithWindow=False)
                prs.SaveAs(new_file_path, FileFormat=new_file_format)
                prs.Close()
                powerpoint.Quit()
                return
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeConvertedError(file, new_file_format)

    @overload
    def copySlide(
        self,
        from_file_path: str,
        slide_num: int,
        new_slide_num: int = -1,
    ) -> None:
        """
        copy_slide(self, from_file_path, slide_num [,new_slide_num=len(self.slides)+1])

        Copy slide from \"from_file_path\" to \"self.file_path\".

        slide_num from 1 to len(self.slides).
        new_slide_num from 1 to len(self.slides)+1.
        """
        t = self.t
        if not os.path.exists(from_file_path):
            raise pptxOpError.FileNotFoundError(from_file_path)
        if not os.path.exists(self.file_path):
            pptxOp.new_file(self.file_path)
        if new_slide_num == -1:
            new_slide_num = len(self.file_path) + 1
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                from_prs = powerpoint.Presentations.Open(
                    from_file_path, WithWindow=False
                )
                from_prs.Slides(slide_num).Copy()
                from_prs.Close()
                to_prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                to_prs.Slides().Paste(new_slide_num)
                to_prs.Save()
                to_prs.Close()
                return
            except Exception as e:
                print(e)
                try:
                    from_prs.Close()
                except Exception:
                    pass
                try:
                    to_prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.SlideCouldNotBeCopiedError(
            from_file_path, slide_num, self.file_path, new_slide_num
        )

    @staticmethod
    def copySlide(
        from_file_path: str,
        slide_num: int,
        new_file_path: str,
        new_slide_num: int = -1,
        t: int = 10,
    ) -> None:
        """
        copy_slide(from_file_path, slide_num, new_file_path,  [,new_slide_num=len(self.slides)+1] [,t=10])

        Copy slide from \"from_file_path\" to \"self.file_path\".

        slide_num from 1 to len(self.slides).
        new_slide_num from 1 to len(self.slides)+1.
        t is the number of repeated trial and error attempts.
        """
        if not os.path.exists(from_file_path):
            raise pptxOpError.FileNotFoundError(from_file_path)
        if not os.path.exists(new_file_path):
            pptxOp.new_file(new_file_path)
        if new_slide_num == -1 or new_slide_num > len(new_file_path):
            new_slide_num = len(new_file_path) + 1
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                from_prs = powerpoint.Presentations.Open(
                    from_file_path, WithWindow=False
                )
                from_prs.Slides(slide_num).Copy()
                from_prs.Close()
                to_prs = powerpoint.Presentations.Open(new_file_path, WithWindow=False)
                to_prs.Slides().Paste(new_slide_num)
                to_prs.Save()
                to_prs.Close()
                return
            except Exception as e:
                print(e)
                try:
                    from_prs.Close()
                except Exception:
                    pass
                try:
                    to_prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.SlideCouldNotBeCopiedError(
            from_file_path, slide_num, new_file_path, new_slide_num
        )

    @overload
    def slidesCount(self) -> int:
        """
        slides_count() -> int

        Return the number of slides in the Powerpoint file.
        """
        t = self.t
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                slides_count = len(prs.Slides)
                prs.Close()
                powerpoint.Quit()
                return slides_count
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeOpenedError(self.file_path)

    @staticmethod
    def slidesCount(file_path: str, t: int = 10) -> int:
        """
        slides_count(file_path [,t=10]) -> int

        Return the number of slides in the Powerpoint file.
        """
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")

                prs = powerpoint.Presentations.Open(file_path, WithWindow=False)
                slides_count = len(prs.Slides)
                prs.Close()
                powerpoint.Quit()
                return slides_count
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.FileCouldNotBeOpenedError(file_path)

    @overload
    def slideSize(self) -> tuple:
        """
        size() -> tuple(Pt, Pt)

        Return the size of the Powerpoint file in (width, height) format.
        """
        t = self.t
        while t > 0:
            try:
                prs = pptx.Presentation(self.file_path)
                width, height = Emu(prs.slide_width).pt, Emu(prs.slide_height).pt
                del prs
                return (width, height)
            except Exception as e:
                print(e)
                try:
                    del prs
                except Exception:
                    pass
                t -= 1
        raise pptxOpError.FileCouldNotBeOpenedError(self.file_path)

    @staticmethod
    def slideSize(file_path: str, t: int = 10) -> tuple:
        """
        size(file_path [,t=10]) -> tuple(Pt, Pt)

        Return the size of the Powerpoint file in (width, height) format.
        """
        while t > 0:
            try:
                prs = pptx.Presentation(file_path)
                width, height = Emu(prs.slide_width).pt, Emu(prs.slide_height).pt
                del prs
                return (width, height)
            except Exception as e:
                print(e)
                try:
                    del prs
                except Exception:
                    pass
                t -= 1
        raise pptxOpError.FileCouldNotBeOpenedError(file_path)
