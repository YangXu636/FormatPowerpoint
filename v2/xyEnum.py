class EnumBase:
    def __getitem__(self, __name):
        return (
            vars(self.__class__)[__name]
            if __name in [i for i in dir(self.__class__) if not i.startswith("_")]
            else None
        )


class PpSaveAsFileType(EnumBase):
    ppSaveAsPresentation = pptx = 1
    ppSaveAsPowerPoint7 = 2
    ppSaveAsPowerPoint4 = 3
    ppSaveAsPowerPoint3 = 4
    ppSaveAsTemplate = 5
    ppSaveAsRTF = rtf = 6
    ppSaveAsShow = 7
    ppSaveAsAddIn = 8
    ppSaveAsPowerPoint4FarEast = 10
    ppSaveAsDefault = default = 11
    ppSaveAsHTML = html = 12
    ppSaveAsHTMLv3 = 13
    ppSaveAsHTMLDual = 14
    ppSaveAsMetaFile = 15
    ppSaveAsGIF = gif = 16
    ppSaveAsJPG = jpg = 17
    ppSaveAsPNG = png = 18
    ppSaveAsBMP = bmp = 19
    ppSaveAsWebArchive = 20
    ppSaveAsTIF = tiff = 21
    ppSaveAsPresForReview = 22
    ppSaveAsEMF = emf = 23
    ppSaveAsOpenXMLPresentation = 24
    ppSaveAsOpenXMLPresentationMacroEnabled = 25
    ppSaveAsOpenXMLTemplate = 26
    ppSaveAsOpenXMLTemplateMacroEnabled = 27
    ppSaveAsOpenXMLShow = 28
    ppSaveAsOpenXMLShowMacroEnabled = 29
    ppSaveAsOpenXMLAddin = 30
    ppSaveAsOpenXMLTheme = 31
    ppSaveAsPDF = pdf = 32
    ppSaveAsXPS = xps = 33
    ppSaveAsXMLPresentation = 34
    ppSaveAsOpenDocumentPresentation = 35
    ppSaveAsOpenXMLPicturePresentation = 36
    ppSaveAsNMV = 37
    ppSaveAsExternalConverter = 64000
