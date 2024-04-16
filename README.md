# FormatPowerpoint
用于格式化ppt

## 运行方式
cmd内运行:
```
python3 main.py
```
打包运行:
```
.venv\Scripts\activate

//生成文件夹
pyinstaller --distpath .\toExe\dist -D main.py

//生成单个文件
pyinstaller --distpath .\toExe\dist -F main.py
```