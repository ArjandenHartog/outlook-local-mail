<#
.SYNOPSIS
  Update outlook-local-mail to the latest release (or a given tag/branch) in one step.
.EXAMPLE
  .\update.ps1                 # latest release tag
  .\update.ps1 -Ref main       # follow the main branch
  .\update.ps1 -Ref v0.3.0
#>
[CmdletBinding()]
param([string]$Ref = 'latest', [switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path $python)) { throw 'Run install.ps1 first.' }
if (git -C $root status --porcelain --untracked-files=no) { throw 'You have local changes to tracked files. Commit or stash them first.' }

git -C $root fetch --tags --quiet
if ($Ref -eq 'latest') {
    $Ref = git -C $root tag --list 'v*' --sort=-v:refname | Select-Object -First 1
    if (-not $Ref) { throw 'No release tags found; use -Ref main.' }
}
Write-Host "==> Updating to $Ref" -ForegroundColor Cyan
if (git -C $root show-ref --verify --quiet "refs/heads/$Ref") { git -C $root checkout --quiet $Ref; git -C $root pull --ff-only --quiet }
elseif (git -C $root show-ref --verify --quiet "refs/remotes/origin/$Ref") { git -C $root checkout --quiet -B $Ref "origin/$Ref" }
else { git -C $root checkout --quiet $Ref }
if ($LASTEXITCODE) { throw "Could not check out $Ref." }

& $python -m pip install --disable-pip-version-check -q -r (Join-Path $root 'bridge\requirements.txt')
if (-not $SkipTests) {
    Push-Location (Join-Path $root 'bridge')
    try {
        $ErrorActionPreference = 'Continue'   # unittest reports on stderr
        $output = & $python -m unittest discover -p 'test_*.py' 2>&1
        $code = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        $output | Select-Object -Last 3 | ForEach-Object { "$_" }
        if ($code) { throw 'Tests failed; the service was not restarted.' }
    } finally { Pop-Location }
}

# Restart the loopback service so it loads the new code (the next mail.py call or the logon task starts it again).
$session = Join-Path $root 'bridge\runtime\session.json'
if (Test-Path $session) {
    $pidValue = (Get-Content $session -Raw | ConvertFrom-Json).pid
    $process = Get-CimInstance Win32_Process -Filter "ProcessId = $([int]$pidValue)" -ErrorAction SilentlyContinue
    if ($process -and $process.Name -match '^pythonw' -and $process.CommandLine -like '*bridge\server.py*') {
        Stop-Process -Id $pidValue
        Write-Host 'Service stopped; it restarts on the next command.'
    }
}
Write-Host "Now at $(git -C $root describe --tags --always)" -ForegroundColor Green
