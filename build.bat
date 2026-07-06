@echo off
chcp 65001 >nul
echo ============================================================
echo  CBT Report — Windows EXE Builder
echo ============================================================
echo.

:: Check uv is available
where uv >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未找到 uv，请先安装：https://docs.astral.sh/uv/getting-started/installation/
    pause
    exit /b 1
)

echo [1/3] 安装依赖（含 pyinstaller）...
uv sync --group dev
if %errorlevel% neq 0 (
    echo [ERROR] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [2/3] 清理上次构建产物...
if exist dist\cbt_report.exe del /f /q dist\cbt_report.exe
if exist build rmdir /s /q build

echo.
echo [3/3] 打包中，请稍候...
uv run pyinstaller cbt_report.spec --clean
if %errorlevel% neq 0 (
    echo [ERROR] 打包失败，请检查上方错误信息
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  打包完成！
echo  输出文件：dist\cbt_report.exe
echo ============================================================
echo.
echo 使用方式：将 cbt_report.exe 与以下文件放在同一目录后双击运行：
echo   - address_db.xlsx
echo   - pod.xlsx
echo   - watch_list.xlsx（可选）
echo.
pause
