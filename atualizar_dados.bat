@echo off
echo ==========================================
echo Sincronizando Notas Fiscais (OneDrive)
echo ==========================================
echo.

:: Agora o script process_data.js faz todo o trabalho de ler a pasta
node process_data.js

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Sucesso! Dashboard atualizado.
) else (
    echo.
    echo Houve um erro no processamento. Verifique se o Node.js esta instalado.
)

echo.
pause
