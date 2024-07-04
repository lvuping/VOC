
$profilePath = $PROFILE

프로파일 파일이 존재하지 않는 경우 파일을 생성합니다:

if (-Not (Test-Path -Path $profilePath)) {
    New-Item -ItemType File -Path $profilePath -Force
}


주어진 함수를 프로파일에 추가합니다:
$functions = @'

notepad $PROFILE