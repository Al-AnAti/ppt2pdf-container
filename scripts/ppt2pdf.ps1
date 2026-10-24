<#
.SYNOPSIS
    Run the ppt2pdf Docker container against a Windows file or folder.

.DESCRIPTION
    Wrapper around `docker run nooralanati/ppt2pdf-container`. Resolves the
    target to an absolute Windows path, mounts the right folder at /workspace
    inside the container, and forwards the converter's flags. PDFs land next to
    their source presentations.

    Add it to your profile (see README "Run it from anywhere on Windows") and
    call it from any directory:

        ppt2pdf                       # everything in the current folder
        ppt2pdf .                     # same thing
        ppt2pdf "D:\Courses\Deck"     # a folder
        ppt2pdf .\lecture.pptx        # one file
        ppt2pdf -r -f                 # recurse and re-convert everything
        ppt2pdf -SplitSteps           # one page per click instead of the final state only
        ppt2pdf -Jobs 2 -Verbose      # two workers, show what was flattened
        ppt2pdf -WhatIf               # print the docker command instead of running it

.PARAMETER Path
    A presentation (.ppt/.pptx and friends) or a directory. Defaults to the
    current directory.
.PARAMETER Recursive
    Also convert presentations in subdirectories (-r).
.PARAMETER Force
    Re-convert even when an up-to-date PDF already exists (-f).
.PARAMETER Jobs
    Number of parallel LibreOffice workers (-j). Default: the container's CPU count.
.PARAMETER Timeout
    Per-file LibreOffice timeout in seconds.
.PARAMETER SplitSteps
    Write one PDF page per animation click (-s) instead of only each slide's final state.
.PARAMETER NoFlatten
    Convert as-is: keep animations and leave hidden shapes hidden.
.PARAMETER Image
    Docker image to run. Defaults to $env:PPT2PDF_IMAGE, then the Docker Hub image.
#>
function ppt2pdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0)]
        [string]$Path = '.',

        [Alias('r')][switch]$Recursive,
        [Alias('f')][switch]$Force,
        [Alias('j')][int]$Jobs,
        [double]$Timeout,
        [Alias('s')][switch]$SplitSteps,
        [switch]$NoFlatten,

        [string]$Image = $(if ($env:PPT2PDF_IMAGE) { $env:PPT2PDF_IMAGE } else { 'nooralanati/ppt2pdf-container:latest' })
    )

    $supported = '.ppt', '.pps', '.pot', '.pptx', '.pptm', '.ppsx', '.potx'
    function Fail([string]$Message) {
        Write-Host "ppt2pdf: $Message" -ForegroundColor Red
        $global:LASTEXITCODE = 2
    }

    # --- 1. Resolve the target to an absolute Windows path -------------------
    try {
        $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    } catch {
        Fail "path not found: '$Path'"
        return
    }
    if ($resolved -like '\\*') {
        Fail "'$resolved' is a UNC path. Docker Desktop cannot mount network shares directly; map it to a drive letter first."
        return
    }

    $item = Get-Item -LiteralPath $resolved -Force
    if ($item.PSIsContainer) {
        $hostDir = $item.FullName
        $target = '.'
    } else {
        if ($supported -notcontains $item.Extension.ToLower()) {
            Fail "'$($item.Name)' is not a PowerPoint file ($($supported -join ', '))"
            return
        }
        $hostDir = $item.DirectoryName
        $target = $item.Name
    }

    # --- 2. Turn C:\Users\me\Decks into C:/Users/me/Decks for the mount ------
    $mountSource = ($hostDir -replace '\\', '/').TrimEnd('/')
    if ($mountSource -match '^[A-Za-z]:$') { $mountSource += '/' }   # drive root, e.g. D:/

    # --- 3. Make sure Docker is there and its daemon is up -------------------
    if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
        Fail "the 'docker' command was not found. Install Docker Desktop (https://www.docker.com/products/docker-desktop/) and reopen your terminal."
        return
    }
    $previousEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $null = & docker version --format '{{.Server.Version}}' 2>&1 } finally { $ErrorActionPreference = $previousEAP }
    if ($LASTEXITCODE -ne 0) {
        Fail "the Docker daemon is not running. Start Docker Desktop, wait until it reports 'Engine running', then try again."
        return
    }

    # --- 4. Build the docker command ------------------------------------------
    $dockerArgs = @('run', '--rm', '--init', '-v', "${mountSource}:/workspace", '-w', '/workspace')
    $interactive = -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected
    if ($interactive) { $dockerArgs += '-it' }       # real TTY: colours, Ctrl-C passes through
    $dockerArgs += $Image

    $toolArgs = @()
    if ($Recursive) { $toolArgs += '--recursive' }
    if ($Force)     { $toolArgs += '--force' }
    if ($Jobs)      { $toolArgs += '--jobs', $Jobs }
    if ($Timeout)   { $toolArgs += '--timeout', $Timeout }
    if ($SplitSteps) { $toolArgs += '--split-steps' }
    if ($NoFlatten)  { $toolArgs += '--no-flatten' }
    if ($PSBoundParameters.ContainsKey('Verbose')) { $toolArgs += '--verbose' }
    $toolArgs += $target

    $display = 'docker ' + (($dockerArgs + $toolArgs) -join ' ')
    if (-not $PSCmdlet.ShouldProcess($hostDir, $display)) {
        Write-Host $display
        return
    }

    & docker @dockerArgs @toolArgs
    # $LASTEXITCODE is left as the container's exit status: 0 = all good, 1 = some file failed.
}
