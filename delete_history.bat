@echo off
del /F /Q "%APPDATA%\Microsoft\Windows\Recent\*"
del /F /Q "%APPDATA%\Microsoft\Windows\Recent\AutomaticDestinations\*"
del /F /Q "%APPDATA%\Microsoft\Windows\Recent\CustomDestinations\*"
echo 최근 문서 기록이 삭제되었습니다.
pause