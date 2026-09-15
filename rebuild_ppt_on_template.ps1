param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string[]]$SourcePaths,

    [string]$Suffix = '-retemplated'
)

$ErrorActionPreference = 'Stop'
$LogPath = Join-Path (Split-Path -Parent $PSCommandPath) 'rebuild_ppt_on_template.log'
Set-Content -Path $LogPath -Value '' -Encoding utf8

function Write-Log {
    param([string]$Message)
    Add-Content -Path $LogPath -Value ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -Encoding utf8
}

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

    throw "Failed to start PowerPoint.Application after retries: $($lastError.Exception.Message)"
}

function Open-Presentation {
    param(
        $Application,
        [string]$Path,
        [bool]$ReadOnly,
        [bool]$Untitled = $false
    )

    $lastError = $null

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $null = $Application.Presentations.Count
            return $Application.Presentations.Open($Path, $ReadOnly, $Untitled, $false)
        } catch {
            $lastError = $_
            Start-Sleep -Seconds 2
        }
    }

    throw "Failed to open presentation '$Path': $($lastError.Exception.Message)"
}

function New-OutputPresentationFromTemplate {
    param(
        $Application,
        [string]$TemplatePath,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    $lastError = $null

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        $presentation = $null
        try {
            $presentation = $Application.Presentations.Add($false)
            $presentation.PageSetup.SlideWidth = $SlideWidth
            $presentation.PageSetup.SlideHeight = $SlideHeight

            for ($i = $presentation.Slides.Count; $i -ge 1; $i--) {
                $slide = $presentation.Slides.Item($i)
                try {
                    $slide.Delete()
                } finally {
                    Release-ComObject -Object $slide
                }
            }

            [void]$presentation.Slides.InsertFromFile($TemplatePath, 0)
            return $presentation
        } catch {
            $lastError = $_
            if ($null -ne $presentation) {
                try {
                    $presentation.Close()
                } catch {
                }
            }
            Release-ComObject -Object $presentation
            Start-Sleep -Seconds 2
        }
    }

    throw "Failed to create output presentation from template '$TemplatePath': $($lastError.Exception.Message)"
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

function Get-SlideText {
    param($Slide)

    $texts = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = $shape.TextFrame.TextRange.Text
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $texts.Add($text)
                }
            }
        } finally {
            Release-ComObject -Object $shape
        }
    }
    return ($texts -join ' ')
}

function Get-BaseSlideIndex {
    param(
        [int]$SlideIndex,
        [string]$SlideText
    )

    if ($SlideIndex -eq 1) {
        return 1
    }

    if ($SlideText -match '(?i)\bq\s*&\s*a\b|\bq\s*and\s*a\b') {
        return 11
    }

    if ($SlideText -match '(?i)thank\s*you|thanks') {
        return 12
    }

    return 5
}

function Remove-TextShapes {
    param($Slide)

    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $shape.Delete()
            }
        } finally {
            Release-ComObject -Object $shape
        }
    }
}

function Should-SkipShape {
    param(
        $Shape,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    try {
        if ($Shape.HasTextFrame) {
            if ($Shape.TextFrame.HasText) {
                return $false
            }

            return $true
        }
    } catch {
    }

    $left = [double]$Shape.Left
    $top = [double]$Shape.Top
    $width = [double]$Shape.Width
    $height = [double]$Shape.Height
    $areaRatio = ($width * $height) / ($SlideWidth * $SlideHeight)

    # Drop source-slide background slabs so the template background can show through.
    if ($areaRatio -ge 0.90 -and $left -le 5000 -and $top -le 5000) {
        return $true
    }

    return $false
}

function Get-CopyableShapeIndices {
    param(
        $Slide,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    $indices = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if (-not (Should-SkipShape -Shape $shape -SlideWidth $SlideWidth -SlideHeight $SlideHeight)) {
                $indices.Add($i) | Out-Null
            }
        } finally {
            Release-ComObject -Object $shape
        }
    }
    return ,$indices.ToArray()
}

function Copy-ShapesToSlide {
    param(
        $SourceSlide,
        $DestSlide,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    $shapeIndices = Get-CopyableShapeIndices -Slide $SourceSlide -SlideWidth $SlideWidth -SlideHeight $SlideHeight

    foreach ($shapeIndex in $shapeIndices) {
        $sourceShape = $null
        $pastedRange = $null
        $lastError = $null

        try {
            $sourceShape = $SourceSlide.Shapes.Item([int]$shapeIndex)

            for ($attempt = 1; $attempt -le 4; $attempt++) {
                try {
                    $sourceShape.Copy()
                    Start-Sleep -Milliseconds 250
                    $pastedRange = $DestSlide.Shapes.Paste()
                    $lastError = $null
                    break
                } catch {
                    $lastError = $_
                    Start-Sleep -Milliseconds 400
                }
            }

            if ($null -ne $lastError) {
                throw "Failed to paste source shape ${shapeIndex}: $($lastError.Exception.Message)"
            }
        } finally {
            Release-ComObject -Object $pastedRange
            Release-ComObject -Object $sourceShape
        }
    }
}

$templateResolved = (Resolve-Path -LiteralPath $TemplatePath).ProviderPath
$sourceResolved = $SourcePaths | ForEach-Object { (Resolve-Path -LiteralPath $_).ProviderPath }

$app = $null

try {
    $app = New-PowerPointApplication
    Write-Log "PowerPoint started."
    $results = @()

    foreach ($sourcePath in $sourceResolved) {
        $sourcePresentation = $null
        $outputPresentation = $null
        $outputPath = Get-OutputPath -SourcePath $sourcePath -SuffixValue $Suffix
        if (Test-Path -LiteralPath $outputPath) {
            Remove-Item -LiteralPath $outputPath -Force
        }

        try {
            Write-Log "Opening source: $sourcePath"
            $sourcePresentation = Open-Presentation -Application $app -Path $sourcePath -ReadOnly $true -Untitled $false
            $slideWidth = [double]$sourcePresentation.PageSetup.SlideWidth
            $slideHeight = [double]$sourcePresentation.PageSetup.SlideHeight
            Write-Log "Creating output from template for: $sourcePath"
            $outputPresentation = New-OutputPresentationFromTemplate -Application $app -TemplatePath $templateResolved -SlideWidth $slideWidth -SlideHeight $slideHeight
            Start-Sleep -Milliseconds 500

            $templateSlideCount = $outputPresentation.Slides.Count

            for ($slideIndex = 1; $slideIndex -le $sourcePresentation.Slides.Count; $slideIndex++) {
                $sourceSlide = $null
                $duplicateRange = $null
                $destSlide = $null

                try {
                    Write-Log "Rebuilding slide $slideIndex / $($sourcePresentation.Slides.Count) from $sourcePath"
                    $sourceSlide = $sourcePresentation.Slides.Item($slideIndex)
                    $slideText = Get-SlideText -Slide $sourceSlide
                    $baseSlideIndex = Get-BaseSlideIndex -SlideIndex $slideIndex -SlideText $slideText
                    $duplicateRange = $outputPresentation.Slides.Item($baseSlideIndex).Duplicate()
                    if ($null -eq $duplicateRange) {
                        throw "Duplicate returned null for base slide $baseSlideIndex."
                    }

                    $destSlide = $duplicateRange.Item(1)
                    if ($null -eq $destSlide) {
                        throw "Duplicate.Item(1) returned null for base slide $baseSlideIndex."
                    }

                    $destSlide.MoveTo($outputPresentation.Slides.Count)
                    Remove-TextShapes -Slide $destSlide

                    Copy-ShapesToSlide -SourceSlide $sourceSlide -DestSlide $destSlide -SlideWidth $slideWidth -SlideHeight $slideHeight
                } catch {
                    throw "Failed while rebuilding slide $slideIndex from '$sourcePath': $($_.Exception.Message)"
                } finally {
                    Release-ComObject -Object $destSlide
                    Release-ComObject -Object $duplicateRange
                    Release-ComObject -Object $sourceSlide
                }
            }

            for ($i = $templateSlideCount; $i -ge 1; $i--) {
                $templateSlide = $outputPresentation.Slides.Item($i)
                try {
                    $templateSlide.Delete()
                } finally {
                    Release-ComObject -Object $templateSlide
                }
            }

            $saveError = $null
            for ($saveAttempt = 1; $saveAttempt -le 4; $saveAttempt++) {
                try {
                    $outputPresentation.SaveAs($outputPath, 24)
                    $saveError = $null
                    break
                } catch {
                    $saveError = $_
                    Start-Sleep -Seconds 2
                }
            }

            if ($null -ne $saveError) {
                throw "Failed to save presentation: $($saveError.Exception.Message)"
            }

            Write-Log "Saved output: $outputPath"

            $results += [PSCustomObject]@{
                Source = $sourcePath
                Output = $outputPath
                Slides = $sourcePresentation.Slides.Count
            }
        } finally {
            if ($null -ne $sourcePresentation) {
                try {
                    $sourcePresentation.Close()
                } catch {
                }
            }

            if ($null -ne $outputPresentation) {
                try {
                    $outputPresentation.Close()
                } catch {
                }
            }

            Release-ComObject -Object $sourcePresentation
            Release-ComObject -Object $outputPresentation
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
