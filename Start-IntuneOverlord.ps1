param(
  [switch]$ForceInstall,
  [switch]$NoOpenBrowser,
  [switch]$WebMode
)

$ErrorActionPreference = 'Stop'

function Test-CommandExists {
  param([Parameter(Mandatory = $true)][string]$Name)

  return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Invoke-NpmCommand {
  param(
    [Parameter(Mandatory = $true)][string[]]$Arguments
  )

  & npm.cmd @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "npm $($Arguments -join ' ') failed with exit code $LASTEXITCODE."
  }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appDirCandidate = Join-Path $scriptRoot 'intune-overlord'

if (Test-Path (Join-Path $appDirCandidate 'package.json')) {
  $appDir = $appDirCandidate
}
elseif (Test-Path (Join-Path $scriptRoot 'package.json')) {
  $appDir = $scriptRoot
}
else {
  throw 'Could not find package.json. Place this script in the repo root or app folder.'
}

if (-not (Test-CommandExists -Name 'node')) {
  $commonNodePaths = @(
    "$env:ProgramFiles\nodejs",
    "${env:ProgramFiles(x86)}\nodejs",
    "$env:LOCALAPPDATA\Programs\nodejs",
    "$env:APPDATA\nvm\$(Get-ChildItem "$env:APPDATA\nvm" -Filter 'v*' -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -ExpandProperty Name -First 1)"
  )

  $foundNodePath = $commonNodePaths | Where-Object { $_ -and (Test-Path (Join-Path $_ 'node.exe')) } | Select-Object -First 1

  if ($foundNodePath) {
    Write-Host "Node.js found at '$foundNodePath' but not in PATH. Adding for this session..." -ForegroundColor Yellow
    $env:PATH = "$foundNodePath;$env:PATH"
  }
  else {
    Write-Host 'Node.js is not installed or is not in PATH.' -ForegroundColor Red
    Write-Host 'Download and install Node.js (LTS) from: https://nodejs.org/en/download' -ForegroundColor Yellow
    throw 'Node.js is required to run Intune Overlord.'
  }
}

if (-not (Test-CommandExists -Name 'npm')) {
  throw 'npm is not installed or is not in PATH.'
}

Push-Location $appDir

try {
  $nodeModulesPath = Join-Path $appDir 'node_modules'
  $nodeBinPath = Join-Path $nodeModulesPath '.bin'
  $tscCmdPath = Join-Path $nodeBinPath 'tsc.cmd'
  $electronCmdPath = Join-Path $nodeBinPath 'electron.cmd'
  $installStampPath = Join-Path $nodeModulesPath '.install-stamp'
  $lockPath = Join-Path $appDir 'package-lock.json'
  $packagePath = Join-Path $appDir 'package.json'
  $distPath = Join-Path $appDir 'dist'
  $distIndexPath = Join-Path $distPath 'index.html'

  $needsInstall = $ForceInstall -or -not (Test-Path $nodeModulesPath) -or -not (Test-Path $tscCmdPath) -or -not (Test-Path $electronCmdPath)

  if (-not $needsInstall) {
    $dependencySourcePath = if (Test-Path $lockPath) { $lockPath } else { $packagePath }

    if (-not (Test-Path $installStampPath)) {
      $needsInstall = $true
    }
    else {
      $stampTime = (Get-Item $installStampPath).LastWriteTimeUtc
      $sourceTime = (Get-Item $dependencySourcePath).LastWriteTimeUtc
      if ($stampTime -lt $sourceTime) {
        $needsInstall = $true
      }
    }
  }

  if ($needsInstall) {
    Write-Host 'Installing dependencies...' -ForegroundColor Cyan
    $installSucceeded = $false

    if (Test-Path $lockPath) {
      try {
        Invoke-NpmCommand -Arguments @('ci')
        $installSucceeded = $true
      }
      catch {
        Write-Host 'npm ci failed, retrying with npm install...' -ForegroundColor Yellow
      }
    }

    if (-not $installSucceeded) {
      Invoke-NpmCommand -Arguments @('install')
    }

    if (-not (Test-Path $nodeModulesPath) -or -not (Test-Path $tscCmdPath) -or -not (Test-Path $electronCmdPath)) {
      throw 'Install step did not create the required local toolchain.'
    }

    New-Item -Path $installStampPath -ItemType File -Force | Out-Null
    Write-Host 'Dependency install complete.' -ForegroundColor Green
  }
  else {
    Write-Host 'Dependencies already installed. Skipping install.' -ForegroundColor Green
  }

  $envPath = Join-Path $appDir '.env'
  $envExamplePath = Join-Path $appDir '.env.example'

  if (-not (Test-Path $envPath) -and (Test-Path $envExamplePath)) {
    Copy-Item $envExamplePath $envPath
    Write-Host 'Created .env from .env.example. Update VITE_AZURE_CLIENT_ID before sign-in.' -ForegroundColor Yellow
  }

  if ($NoOpenBrowser) {
    $env:BROWSER = 'none'
  }

  if ($WebMode) {
    Write-Host 'Starting Intune Overlord in web mode...' -ForegroundColor Cyan
    Invoke-NpmCommand -Arguments @('run', 'dev', '--', '--host')
  }
  else {
    $needsDesktopBuild = $ForceInstall -or -not (Test-Path $distIndexPath)

    if (-not $needsDesktopBuild) {
      $buildInputs = @(
        (Join-Path $appDir 'src'),
        (Join-Path $appDir 'electron'),
        $packagePath,
        (Join-Path $appDir 'vite.config.ts')
      )

      $distTime = (Get-Item $distIndexPath).LastWriteTimeUtc
      foreach ($inputPath in $buildInputs) {
        if (-not (Test-Path $inputPath)) {
          continue
        }

        if ((Get-Item $inputPath).PSIsContainer) {
          $latestInput = Get-ChildItem -Path $inputPath -Recurse -File | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
          if ($latestInput -and $latestInput.LastWriteTimeUtc -gt $distTime) {
            $needsDesktopBuild = $true
            break
          }
        }
        elseif ((Get-Item $inputPath).LastWriteTimeUtc -gt $distTime) {
          $needsDesktopBuild = $true
          break
        }
      }
    }

    if ($needsDesktopBuild) {
      Write-Host 'Building desktop assets...' -ForegroundColor Cyan
      Invoke-NpmCommand -Arguments @('run', 'build')
    }

    Write-Host 'Starting Intune Overlord in desktop mode...' -ForegroundColor Cyan

    if (-not (Test-Path $electronCmdPath)) {
      throw 'Electron launcher was not found after install.'
    }

    Start-Process -FilePath $electronCmdPath -ArgumentList '.' -WorkingDirectory $appDir | Out-Null
    Write-Host 'Desktop app launch requested.' -ForegroundColor Green
  }
}
finally {
  Pop-Location
}
