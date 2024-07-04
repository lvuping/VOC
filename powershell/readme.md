
## 파워쉘 세팅
---

```powershell
$profilePath = $PROFILE
```
프로파일 파일이 존재하지 않는 경우 파일을 생성합니다:

```powershell
if (-Not (Test-Path -Path $profilePath)) {
    New-Item -ItemType File -Path $profilePath -Force
}
```

주어진 함수를 프로파일에 추가합니다:
$functions = @'

notepad $PROFILE

```

---


## Vim 세팅

notepad $PROFILE
```powershell
Set-Alias vim 'C:\Program Files\Neovim\bin\nvim.exe'
```

---
PowerShell에서 Neovim 설정 디렉토리를 수동으로 생성합니다.


```powershell
New-Item -ItemType Directory -Force -Path "$env:LOCALAPPDATA\nvim"
```

init.vim 파일 생성 및 편집

위 명령어가 성공적으로 실행되면, 이제 init.vim 파일을 생성하고 편집할 수 있습니다.

```powershell
notepad "$env:LOCALAPPDATA\nvim\init.vim"
```

init.vim에 기본 설정 추가

```
" Line numbers
set number

" Relative line numbers
set relativenumber

" Syntax highlighting
syntax on

" Enable mouse support
set mouse=a

" Enable clipboard support
set clipboard=unnamedplus

" Set tab width
set tabstop=4
set shiftwidth=4
set expandtab

" Highlight search results
set hlsearch

" Incremental search
set incsearch

" Enable line wrapping
set wrap

" Show command in the last line of the screen
set showcmd

" Show matching parentheses
set showmatch

```

---
