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
        --onedir `
        --windowed `
        --name $AppName `
        --icon ".\assets\app.ico" `
        --distpath ".\dist" `
        --workpath ".\dist\build" `
        --add-data ".\assets;assets" `
        --add-data ".\pyproject.toml;." `
        ".\app.py"

    $TargetDir = Join-Path $Root "dist\$AppName"
    if (-not (Test-Path -LiteralPath (Join-Path $TargetDir "$AppName.exe"))) {
        throw "Build did not produce $TargetDir\$AppName.exe"
    }

    if (Test-Path ".\config.ini") {
        Copy-Item -LiteralPath ".\config.ini" -Destination $TargetDir -Force
    }
    if (Test-Path ".\Original file list.txt") {
        Copy-Item -LiteralPath ".\Original file list.txt" -Destination $TargetDir -Force
    }

    $ZipPath = Join-Path $Root "dist\$AppName.zip"
    if (Test-Path -LiteralPath $ZipPath) {
        Remove-Item -LiteralPath $ZipPath -Force
    }

    $maxRetries = 5
    $retryDelay = 2
    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            Compress-Archive -Path (Join-Path $TargetDir "*") -DestinationPath $ZipPath -Force
            break
        }
        catch {
            if ($i -lt $maxRetries) {
                Write-Host "Compression failed, retrying in $retryDelay seconds... ($i/$maxRetries)"
                Start-Sleep -Seconds $retryDelay
            }
            else {
                throw
            }
        }
    }

    Write-Host "Build complete: $TargetDir"
    Write-Host "Zip complete: $ZipPath"
}
finally {
    Clear-BuildArtifacts
}
