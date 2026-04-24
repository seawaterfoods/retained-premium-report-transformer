@echo off
chcp 65001 >nul

REM ===== 1️⃣ 指定 Java =====
set "JAVA_HOME=C:\Users\user\.jdks\corretto-17.0.18"
set "PATH=%JAVA_HOME%\bin;%PATH%"

echo 使用 Java：
java -version

echo 正在編譯專案...
call mvnw.cmd clean package -DskipTests -q

if %ERRORLEVEL% equ 0 (
    echo 編譯成功！
) else (
    echo 編譯失敗，請檢查錯誤訊息
)

pause