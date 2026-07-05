@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo   부자재 표준견본 등록기 - exe 빌드
echo ============================================
echo.

set "PY=py"
where py >nul 2>&1 || set "PY=python"
where %PY% >nul 2>&1 || (
  echo [ERROR] Python을 찾을 수 없습니다. python.org 에서 설치하세요.
  pause & exit /b 1
)
echo Using interpreter: %PY%
echo.

echo [1/2] 패키지 설치 (인터넷 필요)...
%PY% -m pip install --upgrade pyinstaller pywinauto openpyxl comtypes pywin32 sv-ttk
if errorlevel 1 (
  echo [ERROR] pip install 실패. 인터넷/파이썬 확인.
  pause & exit /b 1
)
echo.

echo [2/2] exe 빌드 (수 분 소요)...
%PY% -m PyInstaller --noconfirm --clean --windowed --name 부자재표준견본등록기 ^
  --uac-admin ^
  --distpath "%~dp0dist" --workpath "%~dp0build" --specpath "%~dp0build" ^
  --collect-all comtypes ^
  --hidden-import comtypes.gen.UIAutomationClient ^
  --hidden-import comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0 ^
  --collect-all pywinauto ^
  --hidden-import sample_export --hidden-import sample_register ^
  --hidden-import customer_map --hidden-import ierp_export ^
  --hidden-import openpyxl --hidden-import sv_ttk ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone ^
  --hidden-import tkinter --hidden-import tkinter.filedialog ^
  registrar_gui.py
echo.

if exist "%~dp0dist\부자재표준견본등록기\부자재표준견본등록기.exe" (
  echo [완료] 배포 폴더 생성됨:
  echo        %~dp0dist\부자재표준견본등록기\
  echo.
  echo  다음: '부자재표준견본등록기' 폴더를 zip으로 압축해 전달.
  echo  받는 사람: 압축 풀고 .exe 더블클릭 ^(대상 PC에 Python 불필요^)
) else (
  echo [실패] exe가 생성되지 않았습니다. 위 오류 메시지를 확인하세요.
)
echo.
pause
endlocal
