import pptxOperationError as pptxOpError
import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa
from pptx.util import Emu
import win32com.client
import os


class pptxOp:
    """Better Powerpoint Operation Library"""

    def __init__(self, file_path: str, t: int = 10) -> None:
        """
        __init__(file_path [,t=10])

        Initialize the pptxOp object.

        Parameters:
            file_path (str): The path of the Powerpoint file.
            t (int): The number of repeated trial and error attempts.

        Returns:
            None
        """
        self.file_path = file_path
        self.t = t

    def newFile(self) -> None:
        """
        newFile()

        Create a new Powerpoint file.

        Parameters:
            None

        Returns:
            None
        """
        t = self.t
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Add()
                if not os.path.exists(os.path.dirname(self.file_path) + "\\"):
                    os.makedirs(os.path.dirname(self.file_path) + "\\")
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

    def pptFileConversion(
        self, new_file_format: str, new_file_path: str = None
    ) -> None:
        """
        pptFileConversion(new_file_format,  [,new_file_path=None])

        Convert Powerpoint file to new_file_format.

        Parameters:
            new_file_format (str): The new file format. 'pptx', 'pdf', 'png', 'jpg', 'gif', 'tiff', 'bmp'
            new_file_path (str): The new file path. If None, the new file will be saved in the same directory as the original file with the same name as the original file but with the new file format.

        Returns:
            None
        """
        t = self.t
        if not os.path.exists(self.file_path):
            raise pptxOpError.FileNotFoundError(self.file_path)
        if new_file_path is None:
            new_file_path = os.path.splitext(self.file_path)[0] + "." + new_file_format
        else:
            new_file_path = (
                new_file_path + os.path.basename(self.file_path).split(".")[0]
            )
        print(new_file_path)
        if not os.path.exists(os.path.dirname(new_file_path)):
            os.makedirs(os.path.dirname(new_file_path))
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                # prs.SaveCopyAs(
                #     new_file_path, FileFormat=PpSaveAsFileType()[new_file_format]
                # )
                prs.SaveCopyAs(new_file_path + "." + new_file_format)
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

    def slideCopy(
        self,
        from_file_path: str,
        slide_num: int,
        new_slide_num: int = -1,
    ) -> None:
        """
        slideCopy(self, from_file_path, slide_num [,new_slide_num=len(self.slides)+1])

        Copy slide from \"from_file_path\" to \"self.file_path\".

        Parameters:
            from_file_path (str): The path of the Powerpoint file to copy from.
            slide_num (int): The number of the slide to copy.
            new_slide_num (int): The number of the new slide. If -1, the new slide will be added to the end of the Powerpoint file.

        Returns:
            None
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

    def slidesCount(self) -> int:
        """
        slidesCount() -> int

        Return the number of slides in the Powerpoint file.

        Parameters:
            None

        Returns:
            int: The number of slides in the Powerpoint file.
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

    def slideSize(self) -> tuple:
        """
        slideSize() -> tuple(Pt, Pt)

        Return the size of the Powerpoint file in (width, height) format.

        Parameters:
            None

        Returns:
            tuple(Pt, Pt): The size of the Powerpoint file in (width, height) format.
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

    def slideTexts(self, slide_num: int) -> list[str]:
        """
        slideText(slide_num) -> list[str]

        Return the text of the slide in slide_num.

        Parameters:
            slide_num (int): The number of the slide to read.

        Returns:
            list[str]: The text of the slide in slide_num.
        """
        t = self.t
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                for i in prs.Slides(slide_num).Shapes:
                    if i.Type == MSO_SHAPE_TYPE.GROUP:
                        i.Ungroup()
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                text = [
                    shape.TextFrame.TextRange.Text
                    for shape in prs.Slides(slide_num).Shapes
                    if shape.TextFrame.HasText
                ]
                prs.Close()
                powerpoint.Quit()
                return text
            except Exception as e:
                print(e)
                try:
                    prs.Close()
                except Exception:
                    pass
                powerpoint.Quit()
                t -= 1
        raise pptxOpError.SlideCouldNotBeReadError(self.file_path, slide_num)

    def slidePictures(self, slide_num: int) -> list:
        """
        slidePictures(slide_num) -> list

        Return the blob of the picture in slide_num.

        Parameters:
            slide_num (int): The number of the slide to read.

        Returns:
            list: The blob of the picture in slide_num.
        """
        t = self.t
        while t > 0:
            try:
                powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
                prs = powerpoint.Presentations.Open(self.file_path, WithWindow=False)
                for i in prs.Slides(slide_num).Shapes:
                    if i.Type == MSO_SHAPE_TYPE.GROUP:
                        i.Ungroup()
                prs.Save()
                prs.Close()
                powerpoint.Quit()
                prs = pptx.Presentation(self.file_path)
                images = [
                    pic.image.blob
                    for pic in prs.slides[slide_num - 1].shapes
                    if pic.shape_type == MSO_SHAPE_TYPE.PICTURE
                ]
                del prs
                return images
            except Exception as e:
                print(e)
                try:
                    del prs
                except Exception:
                    pass
                t -= 1
        raise pptxOpError.SlideCouldNotBeReadError(self.file_path, slide_num)
