# -*- coding:utf-8 -*-
from main_ui import WinGUI as MainWin
import time
import win32com.client

app = MainWin()
if __name__ == "__main__":
    app.log("\n")
    app.success_log("\n")
    app.fail_log("\n")
    app.log(
        f"当前系统时间：{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))}"
    )

    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        powerpoint.Visible = 1
        app.log(f"PowerPoint 版本：{powerpoint.Version}")
        app.log(f"PowerPoint 路径：{powerpoint.Path}")
        if powerpoint.Version < "16.0":
            app.log(
                "PowerPoint 版本可能较低，部分程序可能无法运行，若出现相关报错，请升级到最新版本或至少 2016 版本 。",
                lvl="Warning",
            )
    except Exception as e:
        app.log(
            f"PowerPoint 未安装或未激活，请先安装或激活 PowerPoint。错误信息：{e}",
            lvl="Error",
        )
        powerpoint.Quit()
        time.sleep(60)
        exit()
    app.log(
        f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))} 开始运行"
    )
    app.log(
        "注意！在运行期间请勿手动打开或关闭 PowerPoint ，请勿使用复制、粘贴。否则将导致程序运行异常。"
    )
    app.log("注意！因自识别正确率极低、识别速度慢等原因，v2采用了人工识别。")
    app.log("注意！若模板文件分类错误，请删除sourceFile目录")
    app.protocol("WM_DELETE_WINDOW", app.save_log)
    app.mainloop()
    try:
        powerpoint.Quit()
    except Exception:
        pass
