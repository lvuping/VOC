# 크로스플랫폼 앱 개발 로드맵

## 기술 스택 선정
- 웹: React + Next.js + TailwindCSS
- 모바일: React Native (+ Expo)

## 학습 순서

### 1. React 기본기 (2-3주, 필수)
- React의 핵심 개념 학습
  - 컴포넌트 개념 (기능별 분리, 재사용성, import/export)
  - Props와 State
    - Props: 읽기 전용 외부 데이터 (부모→자식 단방향)
      - Props 자체는 읽기 전용이지만, 부모의 state를 props로 전달하면 값 변경 가능
      - 예시: 부모가 state로 관리하는 값을 자식에게 props로 전달
        - 부모의 setState 함수도 props로 전달하여 자식이 값을 변경할 수 있음
        - 변경된 state는 부모에서 유지되어 다른 컴포넌트에서도 업데이트된 값 사용 가능
  
        ```jsx
        // 부모 컴포넌트
        function ParentComponent() {
          const [count, setCount] = useState(0);
          
          return (
            <div>
              <h1>부모 컴포넌트: {count}</h1>
              <ChildComponent 
                count={count}          // state를 props로 전달
                onIncrement={setCount} // setState 함수를 props로 전달
              />
            </div>
          );
        }

        // 자식 컴포넌트
        function ChildComponent({ count, onIncrement }) {
          return (
            <div>
              <h2>자식 컴포넌트: {count}</h2>
              <button onClick={() => onIncrement(count + 1)}>
                카운트 증가
              </button>
            </div>
          );
        }
        ```
    - State: 컴포넌트 내부 상태 관리
  - Hooks (useState, useEffect 등)
    - useState: 상태 관리를 위한 Hook
      - const [변수, set변수] = useState(초기값)
    - useEffect: 부수 효과 처리를 위한 Hook
      - 컴포넌트 생명주기와 연동
      - 주요 사용 사례:
        - API 데이터 fetching
        - 구독(subscription) 설정
        - DOM 직접 조작
        - 타이머 설정/해제
      - 기본 구조:
        ```jsx
        useEffect(() => {
          // 실행할 부수 효과
          return () => {
            // 정리(cleanup) 함수
          }
        }, [의존성 배열])
        ```
  - 컴포넌트 생명주기
  - 상태 관리
- JavaScript/TypeScript 기본기

### 2. 선택적 심화 학습 (2-3주)
- 상태 관리 (Redux, Recoil 등)
- TypeScript
- 테스팅

### 3. Next.js 또는 React Native (3-4주)
프로젝트 목적에 따라 선택:

#### 웹 서비스가 주력인 경우
React → Next.js → (필요시) React Native

#### 모바일 앱이 주력인 경우
React → React Native → (필요시) Next.js

## 기술별 특징

### Next.js
- SSR (Server-Side Rendering)
- File-based Routing
- API Routes
- SEO 최적화
- 데이터 페칭 방법

### React Native & Expo

#### React Native 특징
- 모바일 특화 컴포넌트
- 모바일 레이아웃 (Flexbox)
- 네이티브 API 활용
- 모바일 특화 UX

#### Expo 특징
- 간편한 개발 환경 설정
- 빠른 프로젝트 시작
- OTA 업데이트 지원
- 제한된 네이티브 모듈 접근

## 주의사항

1. **기본기 중요성**
   - React 기본 개념을 확실히 이해해야 함
   - JavaScript/TypeScript 실력이 중요

2. **실전 프로젝트 중요성**
   - 이론 학습과 실제 구현을 병행
   - 작은 프로젝트부터 시작

3. **공통 개념 vs 플랫폼 특화 개념**
   - 공통 로직은 재사용 가능
   - UI/UX는 플랫폼별로 다르게 접근 필요

4. **UI/UX 차이**
   - TailwindCSS는 React Native에서 NativeWind를 통해 사용 가능
   - 웹과 모바일의 UI/UX 패턴은 다르게 접근 필요

## 개발 시 고려사항

### 장점
1. **코드 재사용성**
   - React 기반으로 비즈니스 로직 공유 가능
   - 컴포넌트 구조와 상태 관리 패턴 유사

2. **개발 효율성**
   - 하나의 개발 팀이 웹과 모바일 모두 개발 가능
   - Next.js의 강력한 웹 최적화 기능

### 주의사항
1. **UI/UX 차이**
   - TailwindCSS는 웹 전용
   - 모바일용 별도 스타일링 필요

2. **코드 구조**
   - 공통 코드와 플랫폼별 코드 분리 필요
   - Monorepo 구조 고려

3. **학습 곡선**
   - 각 플랫폼의 특성 이해 필요
   - 초기 학습에 시간 투자 필요

## 개발 접근 방식 옵션

### 1. 네이티브 앱 개발
- React Native로 완전한 네이티브 앱 개발
- 네이티브한 사용자 경험
- 플랫폼별 기능 최대한 활용 가능

### 2. 하이브리드 (WebView) 방식
- Next.js 웹을 React Native WebView로 래핑
- 장점:
  - 웹 개발 리소스 재사용 가능
  - 빠른 개발 속도
  - 웹 업데이트가 앱 업데이트 없이 가능
- 단점:
  - 네이티브 기능 활용 제한적
    - 푸시 알림 세밀한 제어
    - 백그라운드 작업 처리
    - 하드웨어 센서 직접 접근 (가속도계, 자이로스코프 등)
    - 생체 인증 (Face ID, 지문인식)
    - 파일 시스템 직접 접근
    - 블루투스/NFC 통신
  - 성능이 순수 네이티브 대비 다소 저하
  - 앱스토어 심사시 제약 가능성
