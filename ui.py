import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import os
import time
import control
import asyncio
import win32com.client
from pptx import Presentation  # noqa: F401

mbPPT_Path = ""
needPPT_Path = ""
start_index = 1
finish_index = 2


class WinGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_label_mbPPT_label = self.__tk_label_mbPPT_label(self)
        self.tk_input_mbPPT_input = self.__tk_input_mbPPT_input(self)
        self.tk_button_mbPPT_button = self.__tk_button_mbPPT_button(self)
        self.tk_text_log_text = self.__tk_text_log_text(self)
        global mbPPT_Path
        mbPPT_Path = tk.StringVar()
        global needPPT_Path
        needPPT_Path = tk.StringVar()
        global start_index
        start_index = tk.StringVar()
        global finish_index
        finish_index = tk.StringVar()
        self.tk_text_success_text = self.__tk_text_success_text(self)
        self.tk_text_fail_text = self.__tk_text_fail_text(self)
        self.tk_label_needPPT_label = self.__tk_label_needPPT_label(self)
        self.tk_input_needPPT_input = self.__tk_input_needPPT_input(self)
        self.tk_button_needPPT_button = self.__tk_button_needPPT_button(self)
        self.tk_button_start_button = self.__tk_button_start_button(self)
        self.tk_label_nono1_lable = self.__tk_label_nono1_lable(self)
        self.tk_label_none2_lable = self.__tk_label_none2_lable(self)
        self.tk_label_none3_lable = self.__tk_label_none3_lable(self)
        self.tk_input_start_index_input = self.__tk_input_start_index_input(self)
        self.tk_input_finish_index_input = self.__tk_input_finish_index_input(self)
        self.tk_button_settings_button = self.__tk_button_settings_button(self)
        self.tk_button_about_button = self.__tk_button_about_button(self)
        self.tk_label_fanwei_label = self.__tk_label_fanwei_label(self)
        self.tk_label_to_label = self.__tk_label_to_label(self)
        start_index.trace_add(
            "write",
            lambda a, b, c: self.log(
                f"设置需格式化ppt范围为: {start_index.get()}~{finish_index.get()} {a = } {b = } {c = }"
            ),
        )
        finish_index.trace_add(
            "write",
            lambda a, b, c: self.log(
                f"设置需格式化ppt范围为: {start_index.get()}~{finish_index.get()} {a = } {b = } {c = }"
            ),
        )

    def __win(self):
        self.title("PowerPoint格式化工具")
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

    def __tk_label_mbPPT_label(self, parent):
        label = tk.Label(
            parent,
            text="模板ppt",
            anchor="center",
        )
        label.place(relx=0.0000, rely=0.0000, relwidth=0.1500, relheight=0.0375)
        return label

    def __tk_input_mbPPT_input(self, parent):
        ipt = tk.Entry(
            parent,
            textvariable=mbPPT_Path,
        )
        ipt.place(relx=0.1567, rely=0.0000, relwidth=0.7017, relheight=0.0375)
        return ipt

    def __tk_button_mbPPT_button(self, parent):
        btn = tk.Button(
            parent,
            text="选择模板",
            takefocus=False,
            command=self.selectMbpptPath,
        )
        btn.place(relx=0.8650, rely=0.0000, relwidth=0.1350, relheight=0.0375)
        return btn

    def __tk_text_log_text(self, parent):
        text = tk.Text(
            parent,
            state=tk.DISABLED,
        )
        text.place(relx=0.0000, rely=0.1375, relwidth=0.6200, relheight=0.8625)
        self.create_bar(parent, text, True, True, 0, 55, 372, 345, 600, 400)
        return text

    def __tk_progressbar_main_progressbar(self, parent):
        progressbar = ttk.Progressbar(
            parent,
            orient=tk.HORIZONTAL,
        )
        progressbar.place(relx=0.1583, rely=0.0950, relwidth=0.8367, relheight=0.0400)
        return progressbar

    def __tk_text_success_text(self, parent):
        text = tk.Text(
            parent,
            state=tk.DISABLED,
        )
        text.place(relx=0.6183, rely=0.1375, relwidth=0.1850, relheight=0.8625)
        self.create_bar(parent, text, True, True, 371, 55, 111, 345, 600, 400)
        return text

    def __tk_text_fail_text(self, parent):
        text = tk.Text(
            parent,
            state=tk.DISABLED,
        )
        text.place(relx=0.8033, rely=0.1375, relwidth=0.1967, relheight=0.8600)
        self.create_bar(parent, text, True, True, 482, 55, 118, 344, 600, 400)
        return text

    def __tk_label_needPPT_label(self, parent):
        label = tk.Label(
            parent,
            text="需格式化ppt",
            anchor="center",
        )
        label.place(relx=0.0017, rely=0.0500, relwidth=0.1467, relheight=0.0375)
        return label

    def __tk_input_needPPT_input(self, parent):
        ipt = tk.Entry(
            parent,
            textvariable=needPPT_Path,
        )
        ipt.place(relx=0.1550, rely=0.0500, relwidth=0.7033, relheight=0.0375)
        return ipt

    def __tk_button_needPPT_button(self, parent):
        btn = tk.Button(
            parent,
            text="选择需格",
            takefocus=False,
            command=self.selectNeedpptPath,
        )
        btn.place(relx=0.8633, rely=0.0500, relwidth=0.1333, relheight=0.0375)
        return btn

    def __tk_button_start_button(self, parent):
        btn = tk.Button(
            parent,
            text="开始",
            takefocus=False,
            command=self.startForm,
        )
        btn.place(relx=0.0000, rely=0.0950, relwidth=0.1517, relheight=0.0375)
        return btn

    def __tk_label_nono1_lable(self, parent):
        label = tk.Label(
            parent,
            text="控制台输出",
            anchor="center",
        )
        label.place(relx=0.0000, rely=0.1350, relwidth=0.1217, relheight=0.0375)
        return label

    def __tk_label_none2_lable(self, parent):
        label = tk.Label(
            parent,
            text="成功格式化",
            anchor="center",
        )
        label.place(relx=0.6167, rely=0.1375, relwidth=0.1183, relheight=0.0375)
        return label

    def __tk_label_none3_lable(self, parent):
        label = tk.Label(
            parent,
            text="未成功格式化",
            anchor="center",
        )
        label.place(relx=0.8083, rely=0.1375, relwidth=0.1433, relheight=0.0375)
        return label

    def __tk_input_start_index_input(self, parent):
        ipt = tk.Entry(
            parent,
            textvariable=start_index,
        )
        ipt.place(relx=0.2500, rely=0.0950, relwidth=0.1250, relheight=0.0375)
        return ipt

    def __tk_input_finish_index_input(self, parent):
        ipt = tk.Entry(
            parent,
            textvariable=finish_index,
        )
        ipt.place(relx=0.4000, rely=0.0950, relwidth=0.1250, relheight=0.0375)
        return ipt

    def __tk_button_settings_button(self, parent):
        btn = tk.Button(
            parent,
            text="设置",
            takefocus=False,
        )
        btn.place(relx=0.9117, rely=0.0950, relwidth=0.0833, relheight=0.0375)
        return btn

    def __tk_button_about_button(self, parent):
        btn = tk.Button(
            parent,
            text="关于",
            takefocus=False,
        )
        btn.place(relx=0.8233, rely=0.0950, relwidth=0.0833, relheight=0.0375)
        return btn

    def __tk_label_fanwei_label(self, parent):
        label = tk.Label(
            parent,
            text="范围",
            anchor="center",
        )
        label.place(relx=0.1567, rely=0.0950, relwidth=0.0833, relheight=0.0375)
        return label

    def __tk_label_to_label(self, parent):
        label = tk.Label(
            parent,
            text="~",
            anchor="center",
        )
        label.place(relx=0.3750, rely=0.0950, relwidth=0.0250, relheight=0.0375)
        return label

    def selectMbpptPath(self):
        _path = filedialog.askopenfilename(
            filetypes=[("PowerPoint", ["*.ppt", "*.pptx"])]
        )
        _path = _path.replace("/", "\\")
        global mbPPT_Path
        mbPPT_Path.set(_path)
        self.log(f"设置模板ppt路径为: {mbPPT_Path.get()}")

    def selectNeedpptPath(self):
        _path = filedialog.askopenfilename(
            filetypes=[("PowerPoint", ["*.ppt", "*.pptx"])]
        )
        _path = _path.replace("/", "\\")
        global needPPT_Path
        needPPT_Path.set(_path)
        self.log(f"设置需格式化ppt路径为: {needPPT_Path.get()}")
        global start_index
        start_index.set(1)
        # self.tk_input_start_index_input.config(state="readonly")
        global finish_index
        finish_index.set(control.get_needFormatPpt_count(needPPT_Path.get(), self.log))
        self.log(f"设置需格式化ppt范围为: {start_index.get()}~{finish_index.get()}")
        self.focus_set()

    def startForm(self):
        self.log("加载PowerPoint转换模块...")
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        powerpoint.Visible = 1

        self.log(f"模板ppt路径: {mbPPT_Path.get()}")
        if mbPPT_Path.get() == "":
            self.log("模板ppt路径不能为空。")
            return
        self.log(f"需格式化ppt路径: {needPPT_Path.get()}")
        if needPPT_Path.get() == "":
            self.log("需格式化ppt路径不能为空。")
            return
        self.log("开始复制PPT...")
        localPath = os.getcwd()
        if not os.path.exists(localPath + "\\sourceFile\\"):
            self.log("未找到sourceFile文件夹，正在创建sourceFile文件夹...")
            os.mkdir(localPath + "\\sourceFile\\")
        try:
            control.copy_file(mbPPT_Path.get(), localPath + "\\sourceFile\\")
            control.copy_file(needPPT_Path.get(), localPath + "\\sourceFile\\")
            mbPPT_Name = os.path.basename(mbPPT_Path.get()).split(".")[0]
            needPPT_Name = os.path.basename(needPPT_Path.get()).split(".")[0]
            if mbPPT_Path.get()[-3:] == "ppt":
                ppt1 = powerpoint.Presentations.Open(
                    localPath + "\\sourceFile\\" + mbPPT_Name + ".ppt"
                )
                ppt1.SaveAs(localPath + "\\sourceFile\\" + mbPPT_Name + ".pptx")
                ppt1.Close()
            if needPPT_Path.get()[-3:] == "ppt":
                ppt1 = powerpoint.Presentations.Open(
                    localPath + "\\sourceFile\\" + needPPT_Name + ".ppt"
                )
                ppt1.SaveAs(localPath + "\\sourceFile\\" + needPPT_Name + ".pptx")
                ppt1.Close()
                control.ZipExtract(
                    localPath + "\\sourceFile\\" + needPPT_Name + ".pptx",
                    localPath + "\\sourceFile\\" + needPPT_Name + "\\",
                )
        except Exception as e:
            self.log(f"复制PPT失败, 原因: {e}\n格式化失败。")
            return
        self.log("复制PPT成功。")
        """
        try:
            self.log("复制PPT...")
            control.copy_file(mbPPT_Path.get(), localPath + "\\sourceFile\\")
            mbZIP = (
                localPath
                + "\\sourceFile\\"
                + os.path.basename(mbPPT_Path.get()).split(".")[0]
            )
            if os.path.basename(mbPPT_Path.get())[-3:] == "ppt":
                ppt1 = powerpoint.Presentations.Open(mbZIP + ".ppt")
                ppt1.SaveAs(mbZIP + ".pptx")
                ppt1.Close()
            if os.path.exists(mbZIP + ".zip"):
                os.remove(mbZIP + ".zip")
            os.rename(
                mbZIP + ".pptx",
                mbZIP + ".zip",
            )
            if os.path.exists(mbZIP + "\\"):
                control.remove_file(mbZIP + "\\")
            control.ZipExtract(mbZIP + ".zip", mbZIP + "\\")
            control.copy_file(needPPT_Path.get(), localPath + "\\sourceFile\\")
            needZIP = (
                localPath
                + "\\sourceFile\\"
                + os.path.basename(needPPT_Path.get()).split(".")[0]
            )
            if os.path.basename(needPPT_Path.get())[-3:] == "ppt":
                ppt1 = powerpoint.Presentations.Open(needZIP + ".ppt")
                ppt1.SaveAs(needZIP + ".pptx")
                ppt1.Close()
            if os.path.exists(needZIP + ".zip"):
                os.remove(needZIP + ".zip")
            os.rename(
                needZIP + ".pptx",
                needZIP + ".zip",
            )
            if os.path.exists(needZIP + "\\"):
                control.remove_file(needZIP + "\\")
            control.ZipExtract(needZIP + ".zip", needZIP + "\\")
        except Exception as e:
            self.log(f"复制PPT失败, 原因: {e}\n格式化失败。")
            return
        self.log("复制PPT成功。")
        control.remove_file(mbZIP + ".zip")
        control.remove_file(needZIP + ".zip")
        self.log("开始格式化PPT...")
        # TODO 格式化PPT
        try:
            control.format_ppt(mbZIP + "\\", needZIP + "\\")
        except Exception as e:
            self.log(f"格式化PPT失败, 原因: {e}\n格式化失败。")
            return
        self.log("格式化PPT成功。")
        self.log("开始压缩PPT...")
        try:
            if not os.path.exists(localPath + "\\resultFile\\"):
                self.log("未找到resultFile文件夹，正在创建resultFile文件夹...")
                os.mkdir(localPath + "\\resultFile\\")
            ansZIP = (
                localPath
                + "\\resultFile\\"
                + os.path.basename(needPPT_Path.get()).split(".")[0]
            )
            control.ZipCompress(
                ansZIP + ".zip",
                needZIP + "//",
            )
            os.rename(
                ansZIP + ".zip",
                ansZIP + ".pptx",
            )
        except Exception as e:
            self.log(f"压缩PPT失败, 原因: {e}\n格式化失败。")
            return
        self.log("压缩PPT成功。")
        """
        powerpoint.Quit()
        # try:
        #     control.format_ppt(
        #         localPath + "\\sourceFile\\" + mbPPT_Name + ".pptx",
        #         localPath + "\\sourceFile\\" + needPPT_Name + ".pptx",
        #         localPath
        #         + "\\resultFile\\"
        #         + needPPT_Name
        #         + time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
        #         + ".pptx",
        #         self.log,
        #         self.success_log,
        #         self.fail_log,
        #     )
        # except Exception as e:
        #     self.log(f"格式化PPT失败, 原因: {e}")
        #     return
        global start_index
        global finish_index
        self.thisTime = time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
        control.format_ppt(
            localPath + "\\sourceFile\\" + mbPPT_Name + ".pptx",
            localPath + "\\sourceFile\\" + needPPT_Name + ".pptx",
            localPath + "\\resultFile\\" + needPPT_Name + self.thisTime + ".pptx",
            int(start_index.get()),
            int(finish_index.get()),
            self.log,
            self.success_log,
            self.fail_log,
        )
        self.log("格式化完成。")
        self.log(
            f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))} 格式化完成。"
        )
        self.log("程序结束运行。")
        return

    def save_log(self):
        localPath = os.getcwd()
        if "thisTime" not in vars(self).keys():
            self.thisTime = time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
        if not os.path.exists(localPath + "\\logs\\"):
            os.mkdir(localPath + "\\logs\\")
        os.mkdir(localPath + "\\logs\\" + self.thisTime + "\\")
        try:
            with open(
                localPath + "\\logs\\" + self.thisTime + "\\console.log",
                "w",
                encoding="utf-8",
            ) as f:
                f.write(self.tk_text_log_text.get("1.0", tk.END))
            with open(
                localPath + "\\logs\\" + self.thisTime + "\\success.log",
                "w",
                encoding="utf-8",
            ) as f:
                f.write(self.tk_text_success_text.get("1.0", tk.END))
            with open(
                localPath + "\\logs\\" + self.thisTime + "\\fail.log",
                "w",
                encoding="utf-8",
            ) as f:
                f.write(self.tk_text_fail_text.get("1.0", tk.END))
        except Exception:
            pass
        self.destroy()
        return

    def log(self, msg):
        self.tk_text_log_text.config(state=tk.NORMAL)
        self.tk_text_log_text.insert(tk.END, msg + "\n")
        self.tk_text_log_text.config(state=tk.DISABLED)
        return

    def success_log(self, msg):
        self.tk_text_success_text.config(state=tk.NORMAL)
        self.tk_text_success_text.insert(tk.END, msg + "\n")
        self.tk_text_success_text.config(state=tk.DISABLED)
        return

    def fail_log(self, msg):
        self.tk_text_fail_text.config(state=tk.NORMAL)
        self.tk_text_fail_text.insert(tk.END, msg + "\n")
        self.tk_text_fail_text.config(state=tk.DISABLED)
        return


if __name__ == "__main__":
    win = WinGUI()
    win.log(
        f"\n{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))} 开始运行"
    )
    win.mainloop()
