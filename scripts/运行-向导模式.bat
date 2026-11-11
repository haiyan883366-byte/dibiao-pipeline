@echo off
chcp 65001 >nul
title 底表处理 - 向导模式(C)
setlocal enabledelayedexpansion

echo ============================================
echo   0元课前总览底表处理（向导模式，逐项确认）
echo ============================================
echo.
echo 每一项都显示默认值，直接回车 = 用默认值
echo.

REM ---- 1. 底表路径 ----
echo ① 底表路径
set /p SRC=[直接回车=自动取 D:\download 下日期最新的一份]
echo.

REM ---- 2. 保留学科组名单 ----
echo ② 最新未开课期次的保留学科组名单（逗号分隔，可多个）
set /p KEEP_SU=[默认: 崔国锋,史春晶]
if "!KEEP_SU!"=="" set KEEP_SU=崔国锋,史春晶
echo.

REM ---- 3. 归并目标学科组 ----
echo ③ 归并目标学科组名（1 个）
set /p TARGET_SU=[默认: 张庆波]
if "!TARGET_SU!"=="" set TARGET_SU=张庆波
echo.

REM ---- 4. 未开课期次（一般不用改）----
echo ④ 未开课期次名（一般直接回车，自动取字典序最大）
set /p QICI=[默认: 自动]
echo.

echo ============ 确认 ============
echo 底表路径: !SRC!（空=自动取最新）
echo 保留学科组: !KEEP_SU!
echo 归并目标:   !TARGET_SU!
echo 未开课期次: !QICI!（空=自动）
echo ==============================
choice /C YN /M "确认开始运行？Y=运行 N=退出"
if errorlevel 2 goto :end

set CMD="C:\Users\zhangxian10\.workbuddy\skills\dibiao-pipeline\scripts\dibiao_pipeline.py" --keep-su "!KEEP_SU!" --target-su "!TARGET_SU!"
if not "!SRC!"=="" set CMD=!CMD! --src "!SRC!"
if not "!QICI!"=="" set CMD=!CMD! --qici "!QICI!"

echo.
python !CMD!

echo.
if !errorlevel! neq 0 (
    echo ? 运行出错，请把上面的红字截图发 WorkBuddy 排查
) else (
    echo ? 完成！请打开输出文件，先看「?校验」sheet 的结论行：
    echo    全部绿? = 数据可放心使用
)
:end
echo.
pause
