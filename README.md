# FormatPowerpoint

用于格式化ppt

## 运行方式

cmd内运行:

```none
python3 main.py
```

打包运行:

```none
//激活虚拟环境
.venv\Scripts\activate

//生成文件夹
pyinstaller --distpath .\toExe\dist -D main.py

//生成单个文件
pyinstaller --distpath .\toExe\dist -F main.py
```

## 注意事项

1. 运行前请保存和关闭所有PowerPoint文件和进程。
2. 运行过程中请勿使用PowerPoint 和 关于剪切板的相关操作。
3. 程序运行过程中，请勿关闭cmd窗口。
4. 程序运行结束后若无法使用PowerPoint或打开文件，请使用任务管理器关闭所有PowerPoint进程，并重试。
