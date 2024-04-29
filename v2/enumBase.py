class EnumBase:
    def __getitem__(self, __name):
        return (
            vars(self.__class__)[__name]
            if __name in [i for i in dir(self.__class__) if not i.startswith("_")]
            else None
        )


class PpSaveAsFileType(EnumBase):
    ppSaveAsAddIn = 8
    ppSaveAsAnimatedGIF = 40
    ppSaveAsBMP = bmp = 19
    ppSaveAsDefault = default = pptx = 11
    ppSaveAsEMF = emf = 23
    ppSaveAsExternalConverter = 64000
    ppSaveAsGIF = gif = 16
    ppSaveAsJPG = jpg = 17
    ppSaveAsMetaFile = 15
    ppSaveAsMP4 = mp4 = 39
    ppSaveAsOpenDocumentPresentation = 35
    ppSaveAsOpenXMLAddin = 30
    ppSaveAsOpenXMLPicturePresentation = 36
    ppSaveAsOpenXMLPresentation = 24
    ppSaveAsOpenXMLPresentationMacroEnabled = 25
    ppSaveAsOpenXMLShow = 28
    ppSaveAsOpenXMLShowMacroEnabled = 29
    ppSaveAsOpenXMLTemplate = 26
    ppSaveAsOpenXMLTemplateMacroEnabled = 27
    ppSaveAsOpenXMLTheme = 31
    ppSaveAsPDF = pdf = 32
    ppSaveAsPNG = png = 18
    ppSaveAsPresentation = 1
    ppSaveAsRTF = rtf = 6
    ppSaveAsShow = 7
    ppSaveAsStrictOpenXMLPresentation = 38
    ppSaveAsTemplate = 5
    ppSaveAsTIF = tif = 21
    ppSaveAsWMV = wmv = 37
    ppSaveAsXMLPresentation = 34
    ppSaveAsXPS = xps = 33
