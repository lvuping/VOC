## 개발 환경 설정

### React + Vite 프로젝트 생성 (with pnpm)

1. **pnpm이 설치되어 있지 않다면 설치**
```bash
npm install -g pnpm
```

2. **프로젝트 생성**
```bash
pnpm create vite my-react-app -- --template react
```

TypeScript 사용 시:
```bash
pnpm create vite my-react-app -- --template react-ts
```

3. **프로젝트 설정**
```bash
cd my-react-app    # 프로젝트 폴더로 이동
pnpm install       # 의존성 설치
pnpm run dev           # 개발 서버 실행 (http://localhost:5173)
```

4. **주요 명령어**
```bash
pnpm run dev          # 개발 서버 실행
pnpm run build        # 프로덕션용 빌드
pnpm run preview      # 빌드된 결과물 미리보기
```

### 프로젝트 구조
```
my-react-app/
├── node_modules/     # 프로젝트 의존성 모듈들이 설치되는 폴더
├── public/          # 정적 파일들을 저장하는 폴더 (이미지, 폰트 등)
├── src/             # 소스 코드가 위치하는 메인 폴더
│   ├── assets/      # 프로젝트에서 사용되는 자원 파일들 (이미지, 스타일 등)
│   ├── App.css      # App 컴포넌트의 스타일 파일
│   ├── App.jsx      # 최상위 React 컴포넌트
│   ├── index.css    # 전역 스타일 파일
│   └── main.jsx     # React 애플리케이션의 진입점 (entry point)
├── .gitignore       # Git 버전 관리에서 제외할 파일/폴더 목록
├── index.html       # 애플리케이션의 HTML 템플릿
├── package.json     # 프로젝트 메타데이터 및 의존성 정보
├── vite.config.js   # Vite 설정 파일
└── README.md        # 프로젝트 문서
```



