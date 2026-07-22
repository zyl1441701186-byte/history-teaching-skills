[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [string]$OutputJson = "",

    [double]$TolerancePoints = 1.5
)

$ErrorActionPreference = "Stop"

function Inspect-Shape {
    param(
        $Shape,
        [int]$SlideNumber,
        [System.Collections.Generic.List[object]]$Issues
    )

    $msoGroup = 6
    if ($Shape.Type -eq $msoGroup) {
        for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
            Inspect-Shape -Shape $Shape.GroupItems.Item($i) -SlideNumber $SlideNumber -Issues $Issues
        }
        return
    }

    if ($Shape.HasTextFrame -ne -1) { return }
    if ($Shape.TextFrame2.HasText -ne -1) { return }

    $text = [string]$Shape.TextFrame2.TextRange.Text
    if ([string]::IsNullOrWhiteSpace($text)) { return }

    $frame = $Shape.TextFrame2
    $range = $frame.TextRange
    $availableWidth = [double]$Shape.Width - [double]$frame.MarginLeft - [double]$frame.MarginRight
    $availableHeight = [double]$Shape.Height - [double]$frame.MarginTop - [double]$frame.MarginBottom
    $boundWidth = [double]$range.BoundWidth
    $boundHeight = [double]$range.BoundHeight
    $autoSize = [int]$frame.AutoSize
    $wordWrap = [int]$frame.WordWrap

    $reasons = [System.Collections.Generic.List[string]]::new()
    if ($boundHeight -gt ($availableHeight + $TolerancePoints)) {
        $reasons.Add("height-overflow")
    }
    if ($boundWidth -gt ($availableWidth + $TolerancePoints)) {
        $reasons.Add("width-overflow")
    }
    if ($autoSize -eq 1) {
        $reasons.Add("auto-grow-enabled")
    }
    if ($autoSize -eq 2) {
        $reasons.Add("auto-shrink-enabled")
    }

    if ($reasons.Count -gt 0) {
        $preview = ($text -replace "\s+", " ").Trim()
        if ($preview.Length -gt 100) { $preview = $preview.Substring(0, 100) }
        $Issues.Add([pscustomobject]@{
            slide = $SlideNumber
            shape = [string]$Shape.Name
            reasons = @($reasons)
            availableWidth = [math]::Round($availableWidth, 2)
            availableHeight = [math]::Round($availableHeight, 2)
            boundWidth = [math]::Round($boundWidth, 2)
            boundHeight = [math]::Round($boundHeight, 2)
            autoSize = $autoSize
            wordWrap = $wordWrap
            textPreview = $preview
        })
    }
}

$resolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$powerPoint = $null
$presentation = $null
$issues = [System.Collections.Generic.List[object]]::new()

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($resolvedPptx, $true, $false, $false)

    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            Inspect-Shape -Shape $shape -SlideNumber ([int]$slide.SlideIndex) -Issues $issues
        }
    }

    $result = [pscustomobject]@{
        file = $resolvedPptx
        checkedAt = (Get-Date).ToString("s")
        slideCount = [int]$presentation.Slides.Count
        issueCount = $issues.Count
        passed = ($issues.Count -eq 0)
        issues = @($issues)
    }

    $json = $result | ConvertTo-Json -Depth 8
    if ($OutputJson) {
        $outputDirectory = Split-Path -Parent $OutputJson
        if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        [System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.UTF8Encoding]::new($false))
    }
    $json

    if ($issues.Count -gt 0) { exit 2 }
    exit 0
}
catch {
    [pscustomobject]@{
        file = $PptxPath
        passed = $false
        error = $_.Exception.Message
    } | ConvertTo-Json -Depth 4
    exit 3
}
finally {
    if ($presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($powerPoint) {
        $powerPoint.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
