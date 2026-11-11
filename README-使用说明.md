# 使用说明（换电脑安装）

## 新电脑首次安装（只需一次）

1. 下载本仓库：点右上角绿色 **Code** 按钮 → **Download ZIP**，解压到任意位置
2. 双击 `Install.bat`，按提示走：
   - 如果提示没装 Python：去 https://www.python.org/downloads/ 下载安装，
     **安装时务必勾选 "Add Python to PATH"**，装完再重新双击 `Install.bat`
   - 一切正常会显示 `[OK]` 和 `INSTALL DONE`
3. 桌面会出现 `DailyReport.bat`

## 日常生成日报

把底表 `.xlsx` **拖到桌面 `DailyReport.bat` 图标上**，等跑完自动打开文件所在位置，
生成 `桌面\DailyReport_日期.xlsx`（10 个 sheet，含色阶格式）。

也可以双击 `DailyReport.bat`，粘贴底表路径后回车。

## 更新脚本版本

在这台电脑更新了 skill 并推送 GitHub 后，另一台电脑：
```
git clone https://github.com/haiyan883366-byte/dibiao-pipeline.git
```
或重新 Download ZIP 覆盖，再跑一次 `Install.bat` 即可。

## 底表要求

- Sheet1（或第一个 sheet）为数据表
- 必需字段：期次、基地、学科组名称、账号名称、常驻地城市等级、开课日期、年级名称、辅导姓名、在班人次、转低人次等
- 新版格式需含「加微人次」列（缺失时加微率/删除率自动留空）
