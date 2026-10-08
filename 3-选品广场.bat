@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  银河选品广场 -^> SOP 分档
echo ============================================
echo.
if "%MC_TOKEN%"=="" (
  echo 需要登录态 token。获取方法:
  echo   浏览器登录 https://aiyinhe.art
  echo   F12 -^> Application -^> Local Storage -^> 复制 mc_token
  echo.
  set /p MC_TOKEN="粘贴 token: "
)
echo.
echo [1/2] 全部档位: 24小时加购 ^> 400
py fetch_plaza.py -o "选品广场-SOP分档.xlsx" --min-cart 401
echo.
echo [2/2] 低粉埋伏档: 粉丝 ^< 500 且 加购 1~400
py fetch_plaza.py -o "选品广场-低粉埋伏档.xlsx" --max-followers 500 --min-cart 1 --max-cart 400 --limit 400
echo.
echo 完成，打开两个 xlsx 即可。
pause
