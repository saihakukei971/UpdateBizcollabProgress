@echo off
cd /d "%~dp0\..\Python���s��"

echo [INFO] pip�����������J�n...

:: pip ����
python.exe get-pip.py
if %errorlevel% neq 0 (
    echo [ERROR] pip�̓����Ɏ��s���܂���
    pause
    exit /b 1
)

:: ���C�u�����ꊇ�C���X�g�[��
python.exe -m pip install --upgrade pip
python.exe -m pip install -r "%~dp0\requirements.txt"
if %errorlevel% neq 0 (
    echo [ERROR] ���C�u�����̃C���X�g�[���Ɏ��s���܂���
    pause
    exit /b 1
)

echo [SUCCESS] pip�ƃ��C�u�����̃Z�b�g�A�b�v���������܂���
pause
exit /b 0
