# FormatPowerpoint

用于格式化ppt

## 安装虚拟环境

cmd内运行以下命令：

```none
//请确保您的计算机上以安装过python3.x 、 git等软件并配置好环境变量。推荐使用python3.11版本。

//git clone项目
git clone https://github.com/YangXu636/FormatPowerpoint.git

//进入项目目录
cd FormatPowerpoint

//搭建虚拟环境
python3 -m venv .venv

//激活虚拟环境
.venv\Scripts\activate

//安装依赖
pip install -r requirements.txt
```

## 运行方式

### cmd内运行

```none
python3 main.py
```

### 打包运行

```none
//激活虚拟环境
.venv\Scripts\activate

//生成文件夹
pyinstaller --distpath .\toExe\dist -D main.py
//或生成单个文件
pyinstaller --distpath .\toExe\dist -F main.py
```

## 注意事项

1. 运行前请保存和关闭所有PowerPoint文件和进程。
2. 运行过程中请勿使用PowerPoint 和 关于剪切板的相关操作。
3. 程序运行过程中，非特殊情况请勿关闭cmd窗口。
4. 格式化过程中遇到无法操作PowerPoint文件属于正常情况，这是PowerPoint文件本身性质和win32com库本身不稳定所导致的。出问题时请使用任务管理器关闭程序和PowerPoint进程，并根据第五条进行分布式格式化，最后通过人工合并（如需要程序合并请等待后续版本）
5. 若格式化过程中遇到死循环，请关闭cmd窗口，重新运行程序，并将开始格式化位置从1改成上一次格式化失败的界面的索引值-10。
6. 程序运行结束后若无法使用PowerPoint或打开文件，请使用任务管理器关闭所有PowerPoint进程，并重试。
7. 格式化结束后由于识别精度问题，仍需要人工排查错误。
