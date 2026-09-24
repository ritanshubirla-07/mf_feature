@echo off
REM ==============================================================================
REM AMFI Mutual Fund Daily Ingestion & Dashboard Automation Launcher
REM ==============================================================================

cd /d "%~dp0"
echo Starting AMFI Daily Pipeline...
python run_daily_pipeline.py

if %ERRORLEVEL% EQU 0 (
    echo [SUCCESS] AMFI Daily Pipeline finished successfully.
) else (
    echo [ERROR] AMFI Daily Pipeline encountered an error. Check daily_pipeline.log
)
exit /b %ERRORLEVEL%
