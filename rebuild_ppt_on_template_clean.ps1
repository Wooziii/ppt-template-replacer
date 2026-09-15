param(
    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$AnalysisJsonPath,

    [Parameter(Mandatory = $true)]
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

function New-RgbValue {
    param([int]$Red, [int]$Green, [int]$Blue)
    return $Red + ($Green * 256) + ($Blue * 65536)
}

function Normalize-Text {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }
    return (($Text -replace '\s+', ' ').Trim())
}

function Get-ShapeText {
    param($Shape)
    try {
        if ($Shape.HasTextFrame -and $Shape.TextFrame.HasText) {
            return Normalize-Text -Text $Shape.TextFrame.TextRange.Text
        }
    } catch {
    }
    return ''
}

function Remove-AllSlideShapes {
    param($Slide)

    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        try {
            $shape.Delete()
        } finally {
            Release-ComObject -Object $shape
        }
    }
}

function Test-SkipSourceShape {
    param(
        $Shape,
        [int]$ShapeIndex,
        [int]$TitleShapeIndex,
        [double]$TitleConfidence,
        [string]$Role,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    $text = Get-ShapeText -Shape $Shape

    if ($ShapeIndex -eq $TitleShapeIndex) {
        if ($Role -ne 'content' -or $TitleConfidence -ge 0.75) {
            return $true
        }
    }

    if ($Shape.Name -match 'Slide Number') {
        return $true
    }

    if ($Shape.Type -eq 14 -and $text -match '^\d+$') {
        return $true
    }

    $areaRatio = ([double]$Shape.Width * [double]$Shape.Height) / [math]::Max(1.0, $SlideWidth * $SlideHeight)
    if ($areaRatio -ge 0.90 -and [double]$Shape.Left -le ($SlideWidth * 0.05) -and [double]$Shape.Top -le ($SlideHeight * 0.05)) {
        return $true
    }

    return $false
}

function Get-CopyableShapeCount {
    param(
        $Slide,
        [int]$TitleShapeIndex,
        [double]$TitleConfidence,
        [string]$Role,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    $count = 0
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if (-not (Test-SkipSourceShape -Shape $shape -ShapeIndex $i -TitleShapeIndex $TitleShapeIndex -TitleConfidence $TitleConfidence -Role $Role -SlideWidth $SlideWidth -SlideHeight $SlideHeight)) {
                $count += 1
            }
        } finally {
            Release-ComObject -Object $shape
        }
    }
    return $count
}

function Copy-SourceShapes {
    param(
        $SourceSlide,
        $DestSlide,
        [int]$TitleShapeIndex,
        [double]$TitleConfidence,
        [string]$Role,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    for ($i = 1; $i -le $SourceSlide.Shapes.Count; $i++) {
        $shape = $SourceSlide.Shapes.Item($i)
        $pastedRange = $null
        try {
            if (Test-SkipSourceShape -Shape $shape -ShapeIndex $i -TitleShapeIndex $TitleShapeIndex -TitleConfidence $TitleConfidence -Role $Role -SlideWidth $SlideWidth -SlideHeight $SlideHeight) {
                continue
            }

            $shape.Copy()
            Start-Sleep -Milliseconds 120
            $pastedRange = $DestSlide.Shapes.Paste()
        } finally {
            Release-ComObject -Object $pastedRange
            Release-ComObject -Object $shape
        }
    }
}

function Add-TemplateTextBox {
    param(
        $Slide,
        [string]$Text,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [double]$FontSize,
        [int]$Color,
        [bool]$Bold = $true,
        [int]$Alignment = 1
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $box = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $box.Fill.Visible = 0
    $box.Line.Visible = 0
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.WordWrap = -1
    $range = $box.TextFrame.TextRange
    $range.Text = $Text
    $range.ParagraphFormat.Alignment = $Alignment
    $range.Font.Name = 'Microsoft YaHei'
    try { $range.Font.NameFarEast = 'Microsoft YaHei' } catch {}
    $range.Font.Size = $FontSize
    $range.Font.Bold = if ($Bold) { -1 } else { 0 }
    $range.Font.Color.RGB = $Color
    return $box
}

function Add-Header {
    param($Slide, [string]$Title, [double]$SlideWidth, [double]$SlideHeight)

    $headerLeft = $SlideWidth * 0.064
    $headerTop = $SlideHeight * 0.039
    $headerWidth = $SlideWidth * 0.62
    $headerHeight = $SlideHeight * 0.082
    $fontSize = if ($Title.Length -gt 58) { $SlideHeight * 0.015 } elseif ($Title.Length -gt 34) { $SlideHeight * 0.018 } else { $SlideHeight * 0.023 }
    return Add-TemplateTextBox -Slide $Slide -Text $Title -Left $headerLeft -Top $headerTop -Width $headerWidth -Height $headerHeight -FontSize $fontSize -Color (New-RgbValue 8 25 75) -Bold $true -Alignment 1
}

function Add-CoverText {
    param($Slide, [string]$Title, [string]$Subtitle, [double]$SlideWidth, [double]$SlideHeight)

    $titleFontSize = if ($Title.Length -gt 28) { $SlideHeight * 0.048 } else { $SlideHeight * 0.058 }
    Add-TemplateTextBox -Slide $Slide -Text $Title -Left ($SlideWidth * 0.11) -Top ($SlideHeight * 0.235) -Width ($SlideWidth * 0.78) -Height ($SlideHeight * 0.18) -FontSize $titleFontSize -Color (New-RgbValue 0 0 0) -Bold $true -Alignment 2 | Out-Null
    Add-TemplateTextBox -Slide $Slide -Text $Subtitle -Left ($SlideWidth * 0.25) -Top ($SlideHeight * 0.63) -Width ($SlideWidth * 0.50) -Height ($SlideHeight * 0.07) -FontSize ($SlideHeight * 0.031) -Color (New-RgbValue 0 0 0) -Bold $true -Alignment 2 | Out-Null
}

function Add-SectionTitle {
    param($Slide, [string]$Title, [double]$SlideWidth, [double]$SlideHeight)

    $text = $Title -replace '\s*/\s*', "`r"
    $fontSize = if ($text.Length -gt 46) { $SlideHeight * 0.043 } elseif ($text.Length -gt 26) { $SlideHeight * 0.052 } else { $SlideHeight * 0.066 }
    Add-TemplateTextBox -Slide $Slide -Text $text -Left ($SlideWidth * 0.12) -Top ($SlideHeight * 0.35) -Width ($SlideWidth * 0.76) -Height ($SlideHeight * 0.24) -FontSize $fontSize -Color (New-RgbValue 0 100 206) -Bold $true -Alignment 2 | Out-Null
}

function New-PowerPointApplication {
    $lastError = $null
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            $app = New-Object -ComObject PowerPoint.Application
            $app.Visible = -1
            Start-Sleep -Seconds 2
            $null = $app.Presentations.Count
            return $app
        } catch {
            $lastError = $_
            Start-Sleep -Seconds 2
        }
    }
    throw "Failed to start PowerPoint.Application: $($lastError.Exception.Message)"
}

$templateResolved = (Resolve-Path -LiteralPath $TemplatePath).ProviderPath
$sourceResolved = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
$analysisResolved = (Resolve-Path -LiteralPath $AnalysisJsonPath).ProviderPath
$outputResolved = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $outputResolved
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
if (Test-Path -LiteralPath $outputResolved) {
    Remove-Item -LiteralPath $outputResolved -Force
}

$analysis = (Get-Content -LiteralPath $analysisResolved -Raw -Encoding UTF8 | ConvertFrom-Json).analysis
$slideInfoByIndex = @{}
foreach ($slideInfo in $analysis.slides) {
    $slideInfoByIndex[[int]$slideInfo.slide_index] = $slideInfo
}

$app = $null
$sourcePresentation = $null
$outputPresentation = $null

try {
    $app = New-PowerPointApplication
    $sourcePresentation = $app.Presentations.Open($sourceResolved, $true, $false, $false)
    $slideWidth = [double]$sourcePresentation.PageSetup.SlideWidth
    $slideHeight = [double]$sourcePresentation.PageSetup.SlideHeight

    $outputPresentation = $app.Presentations.Add($false)
    $outputPresentation.PageSetup.SlideWidth = $slideWidth
    $outputPresentation.PageSetup.SlideHeight = $slideHeight
    for ($i = $outputPresentation.Slides.Count; $i -ge 1; $i--) {
        $slide = $outputPresentation.Slides.Item($i)
        try { $slide.Delete() } finally { Release-ComObject -Object $slide }
    }

    [void]$outputPresentation.Slides.InsertFromFile($templateResolved, 0)
    $templateSlideCount = $outputPresentation.Slides.Count

    for ($slideIndex = 1; $slideIndex -le $sourcePresentation.Slides.Count; $slideIndex++) {
        $sourceSlide = $sourcePresentation.Slides.Item($slideIndex)
        $info = $slideInfoByIndex[$slideIndex]
        $title = Normalize-Text -Text $info.title_detection.title
        $titleShapeIndex = [int]$info.title_detection.shape_index
        $titleConfidence = [double]$info.title_detection.confidence
        $role = [string]$info.role
        $copyableCount = Get-CopyableShapeCount -Slide $sourceSlide -TitleShapeIndex $titleShapeIndex -TitleConfidence $titleConfidence -Role $role -SlideWidth $slideWidth -SlideHeight $slideHeight

        if ($slideIndex -eq 1) {
            $baseSlideIndex = 1
            $effectiveRole = 'cover'
        } elseif ($role -eq 'section' -and $copyableCount -eq 0) {
            $baseSlideIndex = 5
            $effectiveRole = 'section'
        } else {
            $baseSlideIndex = 4
            $effectiveRole = 'content'
        }

        $duplicateRange = $outputPresentation.Slides.Item($baseSlideIndex).Duplicate()
        $destSlide = $duplicateRange.Item(1)
        $destSlide.MoveTo($outputPresentation.Slides.Count)
        Remove-AllSlideShapes -Slide $destSlide

        if ($effectiveRole -eq 'cover') {
            $subtitle = ''
            if ($sourceSlide.Shapes.Count -ge 1) {
                $subtitle = Get-ShapeText -Shape $sourceSlide.Shapes.Item(1)
            }
            Add-CoverText -Slide $destSlide -Title $title -Subtitle $subtitle -SlideWidth $slideWidth -SlideHeight $slideHeight
        } elseif ($effectiveRole -eq 'section') {
            Add-Header -Slide $destSlide -Title $title -SlideWidth $slideWidth -SlideHeight $slideHeight | Out-Null
            Add-SectionTitle -Slide $destSlide -Title $title -SlideWidth $slideWidth -SlideHeight $slideHeight
        } else {
            Add-Header -Slide $destSlide -Title $title -SlideWidth $slideWidth -SlideHeight $slideHeight | Out-Null
            Copy-SourceShapes -SourceSlide $sourceSlide -DestSlide $destSlide -TitleShapeIndex $titleShapeIndex -TitleConfidence $titleConfidence -Role $role -SlideWidth $slideWidth -SlideHeight $slideHeight
        }

        Release-ComObject -Object $destSlide
        Release-ComObject -Object $duplicateRange
        Release-ComObject -Object $sourceSlide
    }

    for ($i = $templateSlideCount; $i -ge 1; $i--) {
        $slide = $outputPresentation.Slides.Item($i)
        try { $slide.Delete() } finally { Release-ComObject -Object $slide }
    }

    $outputPresentation.SaveAs($outputResolved, 24)
    [PSCustomObject]@{
        Output = $outputResolved
        Slides = $sourcePresentation.Slides.Count
        Width = $slideWidth
        Height = $slideHeight
    } | Format-List
} finally {
    if ($sourcePresentation -ne $null) {
        try { $sourcePresentation.Close() } catch {}
        Release-ComObject -Object $sourcePresentation
    }
    if ($outputPresentation -ne $null) {
        try { $outputPresentation.Close() } catch {}
        Release-ComObject -Object $outputPresentation
    }
    if ($app -ne $null) {
        try { $app.Quit() } catch {}
        Release-ComObject -Object $app
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
