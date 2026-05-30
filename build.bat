@echo off
chcp 65001 > nul
echo ========================================
echo  결석신고서 자동 생성기 v2 빌드
echo ========================================
echo.

pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo.
echo [빌드 중...]
pyinstaller absence_v2.spec --noconfirm

if errorlevel 1 (
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo.
echo [파일명 변경 중...]
python -c "import os; src=r'dist/absence_v2.exe'; dst=r'dist/결석신고서_생성기_v2.exe'; os.path.exists(dst) and os.remove(dst); os.rename(src, dst); print('완료:', dst)"

echo.
echo ========================================
echo  빌드 완료!  dist\결석신고서_생성기_v2.exe
echo ========================================
pause
