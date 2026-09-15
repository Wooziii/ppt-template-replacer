param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param($Object)

    if ($null -ne $Object) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        } catch {
        }
    }
}

function New-PowerPointApplication {
    $lastError = $null

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $app = New-Object -ComObject PowerPoint.Application
            $app.Visible = -1
            return $app
        } catch {
            $lastError = $_
            Start-Sleep -Seconds 2
        }
    }

    throw "Failed to start PowerPoint.Application after retries: $($lastError.Exception.Message)"
}

$sourceResolved = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [System.IO.Path]::ChangeExtension($sourceResolved, '.pdf')
}

$outputResolved = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $outputResolved
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
}

$app = $null
$presentation = $null

try {
    $app = New-PowerPointApplication
    $presentation = $app.Presentations.Open($sourceResolved, $false, $true, $false)
    Start-Sleep -Milliseconds 500
    $presentation.ExportAsFixedFormat(
        $outputResolved,
        2,
        0,
        $false,
        1,
        1,
        $false,
        $null,
        0,
        $false,
        $false,
        $false,
        $false,
        $false
    )
    $outputResolved
} finally {
    if ($null -ne $presentation) {
        try {
            $presentation.Close()
        } catch {
        }
    }

    if ($null -ne $app) {
        try {
            $app.Quit()
        } catch {
        }
    }

    Release-ComObject -Object $presentation
    Release-ComObject -Object $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
