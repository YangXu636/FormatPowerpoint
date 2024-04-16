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
3. 程序运行过程中，请勿关闭cmd窗口。
4. 程序运行结束后若无法使用PowerPoint或打开文件，请使用任务管理器关闭所有PowerPoint进程，并重试。
