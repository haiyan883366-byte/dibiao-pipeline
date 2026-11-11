@echo off
chcp 65001 >nul
title 底表处理 - 直接模式(A)

echo ============================================
echo   0元课前总览底表处理（直接模式，默认参数）
echo ============================================
echo.

REM ============================================================
REM  【参数修改区】数据参数变化时，只改下面这两行引号里的内容
REM    保留学科组名单：逗号分隔，可写多个
REM    归并目标学科组：只有 1 个名字
REM  （底表路径不用填，自动取 D:\download 下日期最新的一份）
REM ============================================================
set KEEP_SU=崔国锋,史春晶
set TARGET_SU=张庆波
REM ============================================================

python "C:\Users\zhangxian10\.workbuddy\skills\dibiao-pipeline\scripts\dibiao_pipeline.py" --keep-su "%KEEP_SU%" --target-su "%TARGET_SU%"

echo.
if %errorlevel% neq 0 (
    echo ? 运行出错，请把上面的红字截图发 WorkBuddy 排查
) else (
    echo ? 完成！请打开输出文件，先看「?校验」sheet 的结论行：
    echo    全部绿? = 数据可放心使用
)
echo.
pause
