```bash
pip install PyQt5
pip install openpyxl
pip install pywin32
pip install pyautogui
```

Python 스크립트를 EXE 파일로 변환하기 위해 PyInstaller를 사용할 수 있습니다. 다음 단계를 따라주세요:
a. PyInstaller 설치:

```bash
pip install pyinstaller
```

b. 스크립트가 있는 디렉토리로 이동:

```bash
cd path/to/your/script/directory
```

c. PyInstaller를 사용하여 EXE 파일 생성:

```bash
pyinstaller --onefile --windowed your_script_name.py
```

--onefile: 모든 종속성을 포함한 단일 EXE 파일 생성
--windowed: 콘솔 창 없이 GUI 애플리케이션 실행

d. 생성된 EXE 파일은 dist 폴더에서 찾을 수 있습니다.
주의사항:

일부 안티바이러스 프로그램이 PyInstaller로 만든 EXE 파일을 오탐할 수 있습니다. 이는 false positive입니다.
SAP GUI 스크립팅을 사용하는 경우, 최종 사용자의 시스템에 SAP GUI가 설치되어 있어야 하며 스크립팅이 활성화되어 있어야 합니다.
EXE 파일 크기가 큰 편일 수 있습니다. 이는 모든 필요한 Python 환경과 라이브러리가 포함되기 때문입니다.
배포 전 여러 환경에서 테스트를 진행하는 것이 좋습니다.
업데이트가 필요한 경우 전체 EXE 파일을 다시 배포해야 합니다.
