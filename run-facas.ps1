$env:PYTHONDONTWRITEBYTECODE = "1"
python (Join-Path $PSScriptRoot "app.py") @args
