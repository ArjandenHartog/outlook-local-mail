<#
.SYNOPSIS
  One-step setup for outlook-local-mail on Windows.
.DESCRIPTION
  Creates .venv, installs dependencies, detects your Outlook Classic accounts (writes bridge\accounts.json),
  runs the tests, optionally registers a logon task and links the skill into Claude Code / Codex,
  and finishes with a status check. Nothing is sent and no password is read.
.EXAMPLE
  .\install.ps1 -ScheduleTask -LinkSkill both
#>
[CmdletBinding()]
param(
    [switch]$ScheduleTask,                                   # start the background service at logon (Task Scheduler)
    [ValidateSet('none', 'claude', 'codex', 'both')][string]$LinkSkill = 'none',
    [switch]$SkipTests,
    [switch]$SkipAccounts                                    # do not touch Outlook; copy the example config by hand
)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$venv = Join-Path $root '.venv'
$python = Join-Path $venv 'Scripts\python.exe'
$pythonw = Join-Path $venv 'Scripts\pythonw.exe'
$bridge = Join-Path $root 'bridge'

function Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }

Step 'Checking Python (3.11+)'
$launcher = $null
foreach ($candidate in @(@('py', '-3'), @('python'))) {
    if (Get-Command $candidate[0] -ErrorAction SilentlyContinue) {
        $extra = @($candidate | Select-Object -Skip 1)
        $found = & $candidate[0] @extra -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($found -and [version]$found -ge [version]'3.11') { $launcher = $candidate; $version = $found; break }
    }
}
if (-not $launcher) { throw 'Python 3.11 or newer was not found. Install it from python.org or the Microsoft Store and run this again.' }

if (-not (Test-Path $python)) {
    Step "Creating virtual environment (.venv) with Python $version"
    $extra = @($launcher | Select-Object -Skip 1)
    & $launcher[0] @extra -m venv $venv
    if ($LASTEXITCODE) { throw 'Could not create the virtual environment.' }
}
Step 'Installing dependencies'
& $python -m pip install --disable-pip-version-check -q -r (Join-Path $bridge 'requirements.txt')
if ($LASTEXITCODE) { throw 'pip install failed.' }

if (-not $SkipAccounts) {
    Step 'Detecting Outlook Classic accounts (read-only)'
    & $python (Join-Path $bridge 'detect_accounts.py')
    if ($LASTEXITCODE) { Write-Warning 'No accounts detected. Sign in to Outlook Classic and run again, or copy config\accounts.example.json to bridge\accounts.json.' }
} elseif (-not (Test-Path (Join-Path $bridge 'accounts.json'))) {
    Copy-Item (Join-Path $root 'config\accounts.example.json') (Join-Path $bridge 'accounts.json')
    Write-Warning 'Copied the example accounts.json; edit it so it lists your own addresses.'
}

if (-not $SkipTests) {
    Step 'Running tests'
    Push-Location $bridge
    try {
        $ErrorActionPreference = 'Continue'   # unittest reports on stderr
        $output = & $python -m unittest discover -p 'test_*.py' 2>&1
        $code = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        $output | Select-Object -Last 4 | ForEach-Object { "$_" }
        if ($code) { throw 'Tests failed.' }
    } finally { Pop-Location }
}

if ($ScheduleTask) {
    Step 'Registering the logon task "OutlookLocalMail" (restarts the service if it stopped, every 5 minutes)'
    $action = New-ScheduledTaskAction -Execute $pythonw -Argument ('"' + (Join-Path $bridge 'supervisor.py') + '"') -WorkingDirectory $bridge
    $logon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $repeat = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 5)
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 2)
    Register-ScheduledTask -TaskName 'OutlookLocalMail' -Action $action -Trigger @($logon, $repeat) -Settings $settings `
        -Description 'Keeps the outlook-local-mail loopback service running' -Force | Out-Null
}

if ($LinkSkill -ne 'none') {
    $targets = @()
    if ($LinkSkill -in 'claude', 'both') { $targets += (Join-Path $env:USERPROFILE '.claude\skills') }
    if ($LinkSkill -in 'codex', 'both') { $targets += (Join-Path $env:USERPROFILE '.codex\skills') }
    foreach ($folder in $targets) {
        Step "Linking the skill into $folder"
        New-Item -ItemType Directory -Force $folder | Out-Null
        $link = Join-Path $folder 'outlook-local-mail'
        if (Test-Path $link) {
            $item = Get-Item $link -Force
            if ($item.LinkType -eq 'Junction') { Remove-Item $link -Force } else { Write-Warning "$link exists and is not a junction; left untouched."; continue }
        }
        New-Item -ItemType Junction -Path $link -Target (Join-Path $root 'skill') | Out-Null
    }
}

Step 'Status'
& $python (Join-Path $root 'skill\scripts\mail.py') status
Write-Host "`nDone. Use: $python $(Join-Path $root 'skill\scripts\mail.py') status" -ForegroundColor Green
Write-Host 'Optional: add direct IMAP mailboxes with bridge\setup_imap.py (see README).'
