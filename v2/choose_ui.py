import tkinter as tk
import tkinter.ttk as ttk


class WinGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_label_choose_label = self.__tk_label_choose_label(self)
        self.tk_select_box_choose_type_selectbox = (
            self.__tk_select_box_choose_type_selectbox(self)
        )
        self.tk_label_page_label = self.__tk_label_page_label(self)
        self.tk_radio_button_have_image_radiobutton = (
            self.__tk_radio_button_have_image_radiobutton(self)
        )
        self.tk_frame_tails_frame = self.__tk_frame_tails_frame(self)
        self.tk_select_box_settypes_one_selectbox = (
            self.__tk_select_box_settypes_one_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_one_text = self.__tk_text_settypes_one_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_two_selectbox = (
            self.__tk_select_box_settypes_two_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_two_text = self.__tk_text_settypes_two_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_three_selectbox = (
            self.__tk_select_box_settypes_three_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_three_text = self.__tk_text_settypes_three_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_four_selectbox = (
            self.__tk_select_box_settypes_four_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_four_text = self.__tk_text_settypes_four_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_five_selectbox = (
            self.__tk_select_box_settypes_five_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_five_text = self.__tk_text_settypes_five_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_six_selectbox = (
            self.__tk_select_box_settypes_six_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_six_text = self.__tk_text_settypes_six_text(
            self.tk_frame_tails_frame
        )
        self.tk_select_box_settypes_seven_selectbox = (
            self.__tk_select_box_settypes_seven_selectbox(self.tk_frame_tails_frame)
        )
        self.tk_text_settypes_seven_text = self.__tk_text_settypes_seven_text(
            self.tk_frame_tails_frame
        )
        self.tk_button_return = self.__tk_button_return(self)

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

    def __tk_label_choose_label(self, parent):
        label = tk.Label(
            parent,
            text="选择类别",
            anchor="center",
        )
        label.place(relx=0.3333, rely=0.0000, relwidth=0.0833, relheight=0.0750)
        return label

    def __tk_select_box_choose_type_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = (
            "标题",
            "章节",
            "目录",
            "仅图片",
            "仅段落",
            "图片+段落",
            "图片+解释",
            "图片+段落+解释",
            "无用",
        )
        cb.place(relx=0.4167, rely=0.0000, relwidth=0.5000, relheight=0.0750)
        return cb

    def __tk_label_page_label(self, parent):
        label = tk.Label(
            parent,
            text="ppt中第张",
            anchor="center",
        )
        label.place(relx=0.0000, rely=0.0000, relwidth=0.1667, relheight=0.0750)
        return label

    def __tk_radio_button_have_image_radiobutton(self, parent):
        rb = tk.Radiobutton(
            parent,
            text="是否有图片",
        )
        rb.place(relx=0.1667, rely=0.0000, relwidth=0.1667, relheight=0.0750)
        return rb

    def __tk_frame_tails_frame(self, parent):
        frame = tk.Frame(
            parent,
        )
        frame.place(relx=0.0000, rely=0.0775, relwidth=0.9967, relheight=0.8975)
        return frame

    def __tk_select_box_settypes_one_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.0000, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_one_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.0000, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 0, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_two_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.1393, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_two_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.1393, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 50, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_three_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.2786, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_three_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.2786, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 100, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_four_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.4178, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_four_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.4178, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 150, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_five_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.5571, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_five_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.5571, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 200, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_six_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.6964, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_six_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.6964, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 250, 500, 50, 598, 359)
        return text

    def __tk_select_box_settypes_seven_selectbox(self, parent):
        cb = ttk.Combobox(
            parent,
            state="readonly",
        )
        cb["values"] = ("标题", "章节", "目录", "段落", "图片注释", "图片", "无用")
        cb.place(relx=0.8361, rely=0.8357, relwidth=0.1639, relheight=0.1393)
        return cb

    def __tk_text_settypes_seven_text(self, parent):
        text = tk.Text(parent)
        text.place(relx=0.0000, rely=0.8357, relwidth=0.8361, relheight=0.1393)
        self.create_bar(parent, text, True, False, 0, 300, 500, 50, 598, 359)
        return text

    def __tk_button_return(self, parent):
        btn = ttk.Button(
            parent,
            text="确定",
            takefocus=False,
        )
        btn.place(relx=0.9133, rely=0.0000, relwidth=0.0833, relheight=0.0750)
        return btn


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
