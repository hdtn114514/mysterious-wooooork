$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
if (-not (Test-Path '.venv')) { py -m venv .venv }
& .\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
& .\.venv\Scripts\python.exe -m uvicorn backend.main:app --host 127.0.0.1 --port 8878
