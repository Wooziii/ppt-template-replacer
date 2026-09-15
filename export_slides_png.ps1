param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
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

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $app = New-Object -ComObject PowerPoint.Application
            $app.Visible = -1
            Start-Sleep -Seconds 5
            $null = $app.Presentations.Count
            return $app
        } catch {
            $lastError = $_
            Start-Sleep -Seconds 3
        }
    }

    throw "Failed to start PowerPoint.Application: $($lastError.Exception.Message)"
}

function Open-Presentation {
    param(
        $Application,
        [string]$Path
    )

    $lastError = $null

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $null = $Application.Presentations.Count
            return $Application.Presentations.Open($Path, $true, $false, $false)
        } catch {
            $lastError = $_
            Start-Sleep -Seconds 2
        }
    }

    throw "Failed to open presentation '$Path': $($lastError.Exception.Message)"
}

$sourceResolved = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
$outputResolved = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $outputResolved | Out-Null

$app = $null
$presentation = $null

try {
    $app = New-PowerPointApplication
    $presentation = Open-Presentation -Application $app -Path $sourceResolved
    Start-Sleep -Milliseconds 500
    $presentation.SaveAs($outputResolved, 18)
    Get-ChildItem -LiteralPath $outputResolved -Filter '*.png' | Select-Object Name,Length | Sort-Object Name
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
