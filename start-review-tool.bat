@echo off
REM Starts the Course Review Tool on this PC: the app, using the model on this PC.
REM Run it from the project folder (double-click, or type start-review-tool.bat).
set OLLAMA_URL=http://localhost:11434
set OLLAMA_MODEL=qwen3.5:9b
cd /d "%~dp0"
python -m course_review.cli check-model
if errorlevel 1 (
  echo.
  echo The model is not reachable. Start Ollama, check "ollama list" shows qwen3.5:9b, then run this again.
  pause
  exit /b 1
)
python -m course_review.cli serve --port 8080
pause
