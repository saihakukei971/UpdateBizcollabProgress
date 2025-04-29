@echo off
setlocal enabledelayedexpansion

:: ===============================
:: �y1�z���t���擾�i���������O�����ɕύX�j
:: ===============================
for /f %%I in ('powershell -command "(Get-Date).AddDays(-1).ToString('yyyyMMdd')"') do set YYYYMMDD=%%I

:: ===============================
:: �y2�z�e��p�X���`
:: ===============================
set BASE_DIR=C:\Users\rep03\Desktop\Python�t�@�C��\�A�b�v�f�[�g�ƃr�Y�R���{_�i��Report
set PYTHON_EXEC=C:\Users\rep03\Desktop\Python�t�@�C��\Python���s��\python.exe
set LOG_DIR=%BASE_DIR%\log
set LOG_FILE=%LOG_DIR%\�A�b�v�f�[�g�ƃr�Y�R���{�i�����|�[�g���s_%YYYYMMDD%��.log

:: ===============================
:: �y3�z���O�t�H���_�쐬
:: ===============================
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: ===============================
:: �y4�z���O�w�b�_�o��
:: ===============================
echo ==================================================== >> "%LOG_FILE%"
echo [%date% %time%] [INFO] �A�b�v�f�[�g�ƃr�Y�R���{_�������� �J�n >> "%LOG_FILE%"
echo ---------------------------------------------------- >> "%LOG_FILE%"
echo [%date% %time%] [INFO] ���s�f�B���N�g��: %BASE_DIR% >> "%LOG_FILE%"
echo [%date% %time%] [INFO] �g�p����Python: %PYTHON_EXEC% >> "%LOG_FILE%"
echo ==================================================== >> "%LOG_FILE%"

:: ===============================
:: �y5�zPython�̑��݊m�F
:: ===============================
"%PYTHON_EXEC%" --version >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] Python��������܂���B���ݒ���m�F���Ă��������B >> "%LOG_FILE%"
    exit /b 1
)

:: ===============================
:: �y6�zSTEP 01
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 01] �������|�[�g�v���l�擾�ƊeID�i���\�ɔ��f >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\01_�������|�[�g�v���l�擾�ƊeID�i���\�ɔ��f.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 01 ���s���s >> "%LOG_FILE%"
)

:: ===============================
:: �y7�zSTEP 02
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 02] �A�b�v�f�[�g�ƃr�Y�R���{_fam8�i��Report�擾(AX-AD) >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\02_�A�b�v�f�[�g�ƃr�Y�R���{_fam8�i��Report�擾(AX-AD).py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 02 ���s���s >> "%LOG_FILE%"
)

:: ===============================
:: �y8�zSTEP 03
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 03] CSV���W�v�p�V�[�g�iAX-AD�j�փA�b�v���[�h >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\03_CSV���W�v�p�V�[�g�iAX-AD�j�փA�b�v���[�h.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 03 ���s���s >> "%LOG_FILE%"
)

:: ===============================
:: �y9�zSTEP 04
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 04] �����V�[�g��AX-AD�\�Ɋ֐��}���ƒl�̂ݕϊ� >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\04_�����V�[�g��AX-AD�\�Ɋ֐��}���ƒl�̂ݕϊ�.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 04 ���s���s >> "%LOG_FILE%"
)

:: ===============================
:: �y10�zSTEP 05
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 05] �}�C���\_���z�v�Z���ʂ̂݋L�� >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\05_�}�C���\_���z�v�Z���ʂ̂݋L��.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 05 ���s���s >> "%LOG_FILE%"
)

:: ===============================
:: �y11�z��������
:: ===============================
echo ==================================================== >> "%LOG_FILE%"
echo [%date% %time%] [INFO] �S�������� >> "%LOG_FILE%"
echo ==================================================== >> "%LOG_FILE%"

exit /b 0
