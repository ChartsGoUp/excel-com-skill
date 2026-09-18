$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Requirements = Join-Path $ProjectRoot "requirements-windows.txt"
$Worker = Join-Path $PSScriptRoot "office_worker.py"

function Test-PythonCmd($exe, $prefix) {
    try {
        & $exe @prefix --version 2>$null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Resolve-Python {
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) {
        foreach ($tag in @("-3", "-V:3.12", "-V:3.11", "-V:3.13")) {
            if (Test-PythonCmd $py.Source @($tag)) {
                return @($py.Source, $tag)
            }
        }
        # Launcher exists but no tag worked; fall through to direct exes.
    }

    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) {
        if (Test-PythonCmd $python.Source @()) {
            return @($python.Source)
        }
    }

    # Fixed per-user / machine-wide CPython locations (no username hardcoded).
    $candidates = @()
    $candidates += Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    $candidates += Get-ChildItem "C:\Program Files\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    $candidates += Get-ChildItem "C:\Python*\python.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    foreach ($candidate in $candidates) {
        if (Test-PythonCmd $candidate @()) {
            return @($candidate)
        }
    }

    if ($py) {
        # Last resort: return launcher default so the error surfaces clearly.
        return @($py.Source, "-3")
    }

    throw "Windows Python 3 was not found. Install Python for Windows or set OFFICE_COM_PYTHON in WSL."
}

$pythonCmd = Resolve-Python
$exe = $pythonCmd[0]
$prefix = @()
if ($pythonCmd.Length -gt 1) {
    $prefix = $pythonCmd[1..($pythonCmd.Length - 1)]
}

Write-Host "Windows Python: $exe $($prefix -join ' ')"

& $exe @prefix -m pip install --upgrade pip
& $exe @prefix -m pip install -r $Requirements

Write-Host ""
Write-Host "Testing Microsoft Office COM..."
& $exe @prefix $Worker health

if ($LASTEXITCODE -ne 0) {
    throw "Office COM health check failed."
}

Write-Host ""
Write-Host "Office COM bootstrap: PASS"
