# MoonOrange Bulletin PDF Downloader

新浪财经公告PDF批量下载工具，支持股票代码验证、日期范围筛选、预览列表、批量下载与进度追踪。

## 功能特性

- **股票代码验证** - 输入6位A股代码，通过新浪行情接口实时验证并显示股票名称（支持沪市6xxxxx、深市0xxxxx/3xxxxx）
- **日期范围筛选** - 设置起止日期，仅下载所需时间段的公告
- **预览列表** - 下载前预览公告日期、名称及生成文件名，确认无误再下载
- **批量下载** - 一键批量下载已筛选的PDF，自动跳过已存在文件
- **取消下载** - 下载过程中可随时点击取消，剩余任务自动跳过
- **超时设置** - 可配置单文件下载超时秒数（默认60秒），超时自动跳过
- **进度追踪** - 实时进度条 + 带颜色标记的下载日志（成功/失败/跳过/超时/取消）
- **文件名安全处理** - 自动将公告名称中的非法字符（/ \ : * ? " < > |）替换为全角字符，避免保存报错
- **市场自动识别** - 根据股票代码自动判断沪市/深市，拼装正确的PDF下载地址
- **多线程操作** - 验证、获取列表、下载均在后台线程执行，界面不卡顿

## 截图

![UI Preview](screenshot.png)

## 快速开始

### 从源码运行

```bash
pip install requests ttkbootstrap
python MoonOrangeBulletinPDFDownloader.py
```

### 打包EXE

双击 `build.bat` 或手动执行：

```bash
pip install pyinstaller
pyinstaller MoonOrangeBulletinPDFDownloader.spec --clean
```

输出文件位于 `dist/MoonOrangeBulletinPDFDownloader.exe`。

## 使用方法

1. 输入6位股票代码（如 `600388` 龙净环保，`000001` 平安银行）
2. 点击 **验证** 确认股票代码有效
3. 设置日期范围和保存路径
4. 点击 **预览列表** 查看公告列表
5. 点击 **确认下载** 开始批量下载
6. 下载过程中可点击 **取消下载** 中止

## 项目结构

```
MoonOrangeBulletinPDFDownloader.py   # 主程序（GUI + 核心下载逻辑）
MoonOrangeBulletinPDFDownloader.spec # PyInstaller 打包配置
build.bat                            # 一键打包脚本
user_penguin.ico                     # 应用图标
```

## 依赖

- Python 3.8+
- requests
- ttkbootstrap
- tkinter（内置）

## 开源协议

MIT License - Free & Open Source

## 致谢

![UI Preview](user_penguin.png)
