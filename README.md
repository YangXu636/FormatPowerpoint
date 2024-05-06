# FormatPowerpoint

用于格式化ppt

## 安装虚拟环境

cmd内运行以下命令：

```none
//请确保您的计算机上以安装了python3.11 、 git等软件并配置好环境变量。

git clone https://github.com/YangXu636/FormatPowerpoint.git
cd FormatPowerpoint
python3 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 运行方式

### cmd内运行

在`.\FormatPowerpoint`目录下cmd内运行以下命令:

```none
python3 .\v1\main.py
```

### 打包运行

在`.\FormatPowerpoint`目录下cmd内运行以下命令:

```none
//生成文件夹
.venv\Scripts\activate
pyinstaller --distpath .\toExe\dist -D .\v2\main.py

//或生成单个文件
.venv\Scripts\activate
pyinstaller --distpath .\toExe\dist -F .\v2\main.py

//带图标
.venv\Scripts\activate
pyinstaller --distpath .\toExe\dist -F -i icon.ico .\v2\main.py
```

## 注意事项

1. 运行前请保存和关闭所有PowerPoint文件和进程。
2. 运行过程中请勿使用PowerPoint 和 关于剪切板的相关操作。
3. 程序运行过程中，非特殊情况请勿关闭cmd窗口。
4. 格式化过程中遇到无法操作PowerPoint文件属于正常情况，这是PowerPoint文件本身性质和win32com库本身不稳定所导致的。出问题时请使用任务管理器关闭程序和PowerPoint进程，并根据第五条进行分布式格式化，最后通过人工合并（如需要程序合并请等待后续版本）
5. 若格式化过程中遇到死循环，请关闭cmd窗口，重新运行程序，并将开始格式化位置从1改成上一次格式化失败的界面的索引值-10。
6. 程序运行结束后若无法使用PowerPoint或打开文件，请使用任务管理器关闭所有PowerPoint进程，并重试。
7. 格式化结束后由于识别精度问题，仍需要人工排查错误。
8. 程序特异性较为严重，使用其他模板ppt时目前需要按照特殊格式修改模板ppt。
