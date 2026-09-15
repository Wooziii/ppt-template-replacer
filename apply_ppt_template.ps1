param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string[]]$SourcePaths,

    [string]$Suffix = '-模板版'
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

function Get-OutputPath {
    param(
        [string]$SourcePath,
        [string]$SuffixValue
    )

    $directory = Split-Path -Parent $SourcePath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $extension = [System.IO.Path]::GetExtension($SourcePath)
    return Join-Path $directory ($stem + $SuffixValue + $extension)
}

$templateResolved = (Resolve-Path -LiteralPath $TemplatePath).ProviderPath
$sourceResolved = $SourcePaths | ForEach-Object { (Resolve-Path -LiteralPath $_).ProviderPath }

$app = $null

try {
    $app = New-PowerPointApplication
    $results = @()

    foreach ($sourcePath in $sourceResolved) {
        $presentation = $null
        $design = $null
        $outputPath = Get-OutputPath -SourcePath $sourcePath -SuffixValue $Suffix

        try {
            $presentation = $app.Presentations.Open($sourcePath, $false, $false, $false)
            Start-Sleep -Milliseconds 500

            $presentation.ApplyTemplate($templateResolved)
            $design = $presentation.Designs.Item(1)

            foreach ($slide in $presentation.Slides) {
                $slide.Design = $design
                Release-ComObject -Object $slide
            }

            $presentation.SaveCopyAs($outputPath)

            $results += [PSCustomObject]@{
                Source = $sourcePath
                Output = $outputPath
                Slides = $presentation.Slides.Count
                Width = $presentation.PageSetup.SlideWidth
                Height = $presentation.PageSetup.SlideHeight
            }
        } finally {
            if ($null -ne $presentation) {
                try {
                    $presentation.Close()
                } catch {
                }
            }

            Release-ComObject -Object $design
            Release-ComObject -Object $presentation
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 1
        }
    }

    $results | Format-Table -AutoSize
} finally {
    if ($null -ne $app) {
        try {
            $app.Quit()
        } catch {
        }
    }

    Release-ComObject -Object $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
