# PowerShell 스크립트로 최근 문서 기록 삭제

# 안전한 파일 삭제를 위해 -Force 매개변수 사용
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Recent\*" -Force -Recurse
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\*" -Force -Recurse
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations\*" -Force -Recurse

Write-Output "최근 문서 기록이 삭제되었습니다."

# 사용자 입력 대기 (pause 명령어와 유사한 기능)
Read-Host -Prompt "계속하려면 Enter 키를 누르세요..."
