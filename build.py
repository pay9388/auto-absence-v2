"""EXE 빌드 스크립트"""
import os
import subprocess
import sys


def run(cmd):
    result = subprocess.run(cmd, shell=True)
    if result.returncode != 0:
        print(f'[오류] 명령 실패: {cmd}')
        sys.exit(1)


print('=' * 40)
print(' 결석신고서 자동 생성기 v2 빌드')
print('=' * 40)

run('pip install -r requirements.txt')

print('\n[빌드 중...]')
run('pyinstaller absence_v2.spec --noconfirm')

src = r'dist/absence_v2.exe'
dst = r'dist/결석신고서_생성기_v2.exe'
if os.path.exists(dst):
    os.remove(dst)
os.rename(src, dst)

print(f'\n완료: {dst}')
print('=' * 40)
