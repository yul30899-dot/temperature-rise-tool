@echo off
chcp 65001 >nul
color 0A
echo =========================================
echo       一键上传代码到 GitHub 工具
echo =========================================
echo.

cd /d "%~dp0"

echo [1/3] 正在扫描修改过的文件...
git add .

set /p msg="请输入这次更新的说明文字 (不输直接回车则默认使用当前时间): "
if "%msg%"=="" (
    set msg=Auto update %date% %time:~0,8%
)

echo.
echo [2/3] 正在打标签: [%msg%]
git commit -m "%msg%"

echo.
echo [3/3] 正在向 GitHub 云端发射...
git push

echo.
if %errorlevel% equ 0 (
    echo =========================================
    echo          上传成功！代码已备份至云端！
    echo =========================================
) else (
    color 0C
    echo =========================================
    echo     上传遇到了一点小麻烦，请看上面的报错提示
    echo =========================================
)

echo.
pause
