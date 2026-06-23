$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$AppName = "3d-batch-copy"

function Remove-InWorkspace {
    param([string]$RelativePath)

    $target = Join-Path $Root $RelativePath
    if (-not (Test-Path -LiteralPath $target)) {
        return
    }

    $resolvedRoot = (Resolve-Path -LiteralPath $Root).Path
    $resolvedTarget = (Resolve-Path -LiteralPath $target).Path
    if (-not $resolvedTarget.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove outside workspace: $resolvedTarget"
    }

    Remove-Item -LiteralPath $resolvedTarget -Recurse -Force
}

function Clear-BuildArtifacts {
    Remove-InWorkspace ".venv"
    Remove-InWorkspace "build"
    Remove-InWorkspace ".pytest_cache"
    Remove-InWorkspace "$AppName.spec"
    Remove-InWorkspace "dist\build"
    Remove-InWorkspace "dist\$AppName"

    Get-ChildItem -LiteralPath $Root -Directory -Recurse -Force -Filter "__pycache__" |
        Where-Object { $_.FullName.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase) } |
        Remove-Item -Recurse -Force
}

try {
    Clear-BuildArtifacts

    python -m venv .venv

    & ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
    & ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt

    if (-not (Test-Path ".\assets\app.ico")) {
        throw "Missing icon file: assets\app.ico"
    }

    & ".\.venv\Scripts\pyinstaller.exe" `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name $AppName `
        --icon ".\assets\app.ico" `
        --distpath ".\dist" `
        --workpath ".\dist\build" `
        --add-data ".\assets;assets" `
        --add-data ".\pyproject.toml;." `
        ".\app.py"

    $ExePath = Join-Path $Root "dist\$AppName.exe"
    if (-not (Test-Path -LiteralPath $ExePath)) {
        throw "Build did not produce $ExePath"
    }

    Write-Host "Build complete: $ExePath"
}
finally {
    Clear-BuildArtifacts
}
