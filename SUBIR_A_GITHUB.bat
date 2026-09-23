@echo off
chcp 65001 >nul
echo ========================================================
echo   CONQUISTA TU ZONA - SUBIR PROYECTO A GITHUB / VERCEL
echo ========================================================
echo.

set "GIT_EXE=git"
where git >nul 2>nul
if %errorlevel% neq 0 (
    if exist "C:\Users\analista.retail\AppData\Local\Programs\Git\cmd\git.exe" (
        set "GIT_EXE=C:\Users\analista.retail\AppData\Local\Programs\Git\cmd\git.exe"
    ) else (
        echo [ERROR] No se encontro Git instalado en el sistema.
        pause
        exit /b 1
    )
)

echo Usando Git: %GIT_EXE%
echo.

if not exist ".git" (
    echo [INFO] Inicializando repositorio Git...
    "%GIT_EXE%" init
    "%GIT_EXE%" branch -M main
)

echo [INFO] Agregando archivos al control de versiones...
"%GIT_EXE%" add .

set /p MSG="Ingresa una descripcion del cambio (o presiona ENTER para automatica): "
if "%MSG%"=="" set MSG="Actualizacion Escuadron de Combate Quest"

"%GIT_EXE%" commit -m "%MSG%"

echo.
"%GIT_EXE%" remote get-url origin >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ========================================================
    echo  CONFIGURACION DEL REPOSITORIO REMOTO
    echo ========================================================
    echo Ingresa la URL de tu repositorio en GitHub
    echo Ejemplo: https://github.com/Santozsj61/quest-escuadron.git
    echo.
    set /p REPO_URL="URL del repositorio: "
    if not "%REPO_URL%"=="" (
        "%GIT_EXE%" remote add origin %REPO_URL%
        echo [OK] Remote agregado.
    ) else (
        echo [AVISO] No ingresaste URL de repositorio remoto. Se guardo localmente.
        pause
        exit /b 0
    )
)

echo.
echo [INFO] Subiendo cambios a GitHub...
"%GIT_EXE%" push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ========================================================
    echo  SUBIDA EXITOSA A GITHUB!
    echo  Vercel detectara el cambio y actualizara el sitio en segundos.
    echo ========================================================
) else (
    echo.
    echo [ERROR] Hubo un problema al subir a GitHub. Revisa tus credenciales o conexion.
)

echo.
pause
