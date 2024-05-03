import tkinter as tk
from PIL import Image, ImageTk

chooseText = ""
texts = ""
image = None


def getType(**kwargs):
    """
    选择框

    Parameters:
        text(str): 文本内容
        image(_io.BytesIO): 图片blob数据

    Returns:
        str: 选择的文本内容
    """
    global texts, image
    texts = kwargs.get("text", "")
    image = kwargs.get("image", None)
    root = kwargs.get("root", None)
    if not root:
        raise ValueError("root参数不能为空")
    if image:
        image = Image.open(image)
        image = image.resize((600, 300))
        image = ImageTk.PhotoImage(image)
    chooseUi = WinGUI()
    root.wait_window(chooseUi)
    global chooseText
    return chooseText if chooseText else None


class WinGUI(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_text_word = self.__tk_text_word(self)
        self.tk_button_chooseTitle = self.__tk_button_chooseTitle(self)
        self.tk_button_chooseChapter = self.__tk_button_chooseChapter(self)
        self.tk_button_chooseCatalogue = self.__tk_button_chooseCatalogue(self)
        self.tk_button_chooseParagraph = self.__tk_button_chooseParagraph(self)
        self.tk_button_chooseExplain = self.__tk_button_chooseExplain(self)
        self.tk_button_chooseImage = self.__tk_button_chooseImage(self)
        self.tk_button_chooseNoneed = self.__tk_button_chooseNoneed(self)
        self.tk_button_chooseOnlyParagraph = self.__tk_button_chooseOnlyParagraph(self)
        self.tk_button_chooseOnlyImage = self.__tk_button_chooseOnlyImage(self)
        self.tk_button_chooseImageneirong = self.__tk_button_chooseImageneirong(self)
        self.tk_button_chooseImageAndParagraph = (
            self.__tk_button_chooseImageAndParagraph(self)
        )
        self.tk_button_chooseImageAndExplain = self.__tk_button_chooseImageAndExplain(
            self
        )
        global image
        if image:
            self.tk_text_word.image_create("1.0", image=image)

    def __win(self):
        self.title("选择")
        # 设置窗口大小、居中
        width = 600
        height = 400
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = "%dx%d+%d+%d" % (
            width,
            height,
            (screenwidth - width) / 2,
            (screenheight - height) / 2,
        )
        self.geometry(geometry)
        self.minsize(width=width, height=height)

    def scrollbar_autohide(self, vbar, hbar, widget):
        """自动隐藏滚动条"""

        def show():
            if vbar:
                vbar.lift(widget)
            if hbar:
                hbar.lift(widget)

        def hide():
            if vbar:
                vbar.lower(widget)
            if hbar:
                hbar.lower(widget)

        hide()
        widget.bind("<Enter>", lambda e: show())
        if vbar:
            vbar.bind("<Enter>", lambda e: show())
        if vbar:
            vbar.bind("<Leave>", lambda e: hide())
        if hbar:
            hbar.bind("<Enter>", lambda e: show())
        if hbar:
            hbar.bind("<Leave>", lambda e: hide())
        widget.bind("<Leave>", lambda e: hide())

    def v_scrollbar(self, vbar, widget, x, y, w, h, pw, ph):
        widget.configure(yscrollcommand=vbar.set)
        vbar.config(command=widget.yview)
        vbar.place(relx=(w + x) / pw, rely=y / ph, relheight=h / ph, anchor="ne")

    def h_scrollbar(self, hbar, widget, x, y, w, h, pw, ph):
        widget.configure(xscrollcommand=hbar.set)
        hbar.config(command=widget.xview)
        hbar.place(relx=x / pw, rely=(y + h) / ph, relwidth=w / pw, anchor="sw")

    def create_bar(self, master, widget, is_vbar, is_hbar, x, y, w, h, pw, ph):
        vbar, hbar = None, None
        if is_vbar:
            vbar = tk.Scrollbar(master)
            self.v_scrollbar(vbar, widget, x, y, w, h, pw, ph)
        if is_hbar:
            hbar = tk.Scrollbar(master, orient="horizontal")
            self.h_scrollbar(hbar, widget, x, y, w, h, pw, ph)
        self.scrollbar_autohide(vbar, hbar, widget)

    def return_choose(self, text):
        global chooseText
        chooseText = text
        self.destroy()

    def __tk_text_word(self, parent):
        global texts
        text = tk.Text(parent)
        text.insert(tk.INSERT, texts)
        text.config(state=tk.DISABLED)
        text.place(relx=0.0000, rely=0.0000, relwidth=1.0000, relheight=0.7500)
        self.create_bar(parent, text, True, True, 0, 0, 600, 300, 600, 400)
        return text

    def __tk_button_chooseTitle(self, parent):
        btn = tk.Button(
            parent,
            text="标题",
            takefocus=False,
            command=lambda: self.return_choose("标题"),
        )
        btn.place(relx=0.0000, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseChapter(self, parent):
        btn = tk.Button(
            parent,
            text="章节",
            takefocus=False,
            command=lambda: self.return_choose("章节"),
        )
        btn.place(relx=0.0833, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseCatalogue(self, parent):
        btn = tk.Button(
            parent,
            text="目录",
            takefocus=False,
            command=lambda: self.return_choose("目录"),
        )
        btn.place(relx=0.1667, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseParagraph(self, parent):
        btn = tk.Button(
            parent,
            text="段落",
            takefocus=False,
            command=lambda: self.return_choose("段落"),
        )
        btn.place(relx=0.2500, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseExplain(self, parent):
        btn = tk.Button(
            parent,
            text="图片注释",
            takefocus=False,
            command=lambda: self.return_choose("图片注释"),
        )
        btn.place(relx=0.3333, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseImage(self, parent):
        btn = tk.Button(
            parent,
            text="图片",
            takefocus=False,
            command=lambda: self.return_choose("图片"),
        )
        btn.place(relx=0.4167, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseNoneed(self, parent):
        btn = tk.Button(
            parent,
            text="无用",
            takefocus=False,
            command=lambda: self.return_choose("无用"),
        )
        btn.place(relx=0.9167, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseOnlyParagraph(self, parent):
        btn = tk.Button(
            parent,
            text="仅段落",
            takefocus=False,
            command=lambda: self.return_choose("仅段落"),
        )
        btn.place(relx=0.5833, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseOnlyImage(self, parent):
        btn = tk.Button(
            parent,
            text="仅图片",
            takefocus=False,
            command=lambda: self.return_choose("仅图片"),
        )
        btn.place(relx=0.6667, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseImageneirong(self, parent):
        btn = tk.Button(
            parent,
            text="图片角标",
            takefocus=False,
            command=lambda: self.return_choose("图角标"),
        )
        btn.place(relx=0.5000, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseImageAndParagraph(self, parent):
        btn = tk.Button(
            parent,
            text="图+段",
            takefocus=False,
            command=lambda: self.return_choose("图+段"),
        )
        btn.place(relx=0.7467, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn

    def __tk_button_chooseImageAndExplain(self, parent):
        btn = tk.Button(
            parent,
            text="图+释",
            takefocus=False,
            command=lambda: self.return_choose("图+释"),
        )
        btn.place(relx=0.8333, rely=0.7500, relwidth=0.0833, relheight=0.2500)
        return btn


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
