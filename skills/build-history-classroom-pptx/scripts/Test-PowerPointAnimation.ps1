[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkDir,

    [switch]$KeepArtifact
)

$ErrorActionPreference = 'Stop'
$ppt = $null
$pres = $null
$reopened = $null
$slide = $null
$shape1 = $null
$shape2 = $null
$testPath = $null

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

try {
    $resolvedWorkDir = [System.IO.Path]::GetFullPath($WorkDir)
    if (-not (Test-Path -LiteralPath $resolvedWorkDir)) {
        [void](New-Item -ItemType Directory -Path $resolvedWorkDir)
    }

    $testPath = Join-Path $resolvedWorkDir ("powerpoint-animation-preflight-{0}.pptx" -f [guid]::NewGuid().ToString('N'))

    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Add()
    $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

    $shape1 = $slide.Shapes.AddTextbox(1, 72, 72, 320, 60)
    $shape1.Name = 'Probe_Question'
    $shape1.TextFrame.TextRange.Text = '问题'

    $shape2 = $slide.Shapes.AddTextbox(1, 72, 160, 320, 60)
    $shape2.Name = 'Probe_Answer'
    $shape2.TextFrame.TextRange.Text = '答案'

    [void]$slide.TimeLine.MainSequence.AddEffect($shape1, 1, 0, 1, -1)  # appear, onClick
    [void]$slide.TimeLine.MainSequence.AddEffect($shape2, 10, 0, 2, -1) # fade, withPrevious

    $pres.SaveAs($testPath, 24) # ppSaveAsOpenXMLPresentation
    $pres.Close()
    Release-ComObject $pres
    $pres = $null

    $reopened = $ppt.Presentations.Open($testPath, -1, 0, 0)
    $sequenceCount = $reopened.Slides.Item(1).TimeLine.MainSequence.Count
    $shapeNames = @(
        $reopened.Slides.Item(1).TimeLine.MainSequence.Item(1).Shape.Name
        $reopened.Slides.Item(1).TimeLine.MainSequence.Item(2).Shape.Name
    )

    if ($sequenceCount -ne 2) {
        throw "Animation sequence count was $sequenceCount instead of 2."
    }
    if ($shapeNames[0] -ne 'Probe_Question' -or $shapeNames[1] -ne 'Probe_Answer') {
        throw 'Animation sequence shape names changed after reopening.'
    }

    $reopened.Close()
    Release-ComObject $reopened
    $reopened = $null

    if (-not $KeepArtifact -and (Test-Path -LiteralPath $testPath)) {
        Remove-Item -LiteralPath $testPath -Force
    }

    [pscustomobject]@{
        status = 'passed'
        powerpointAutomation = $true
        nativeAnimation = $true
        sequenceCount = $sequenceCount
        artifact = if ($KeepArtifact) { $testPath } else { $null }
    } | ConvertTo-Json -Depth 4
    exit 0
}
catch {
    [pscustomobject]@{
        status = 'blocked'
        powerpointAutomation = $false
        nativeAnimation = $false
        error = $_.Exception.Message
        artifact = $testPath
    } | ConvertTo-Json -Depth 4
    exit 1
}
finally {
    if ($null -ne $reopened) {
        try { $reopened.Close() } catch {}
    }
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
    }
    Release-ComObject $shape2
    Release-ComObject $shape1
    Release-ComObject $slide
    Release-ComObject $reopened
    Release-ComObject $pres
    Release-ComObject $ppt
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
