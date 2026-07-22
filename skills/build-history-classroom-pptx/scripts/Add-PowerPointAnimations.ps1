[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPptx,

    [Parameter(Mandatory = $true)]
    [string]$PlanPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPptx,

    [switch]$Overwrite,

    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'

$effectMap = @{
    appear = 1
    fade = 10
    wipe = 22
}

$triggerMap = @{
    onClick = 1
    withPrevious = 2
    afterPrevious = 3
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Validate-Plan {
    param([object]$Plan)

    if ($null -eq $Plan.slides) {
        throw 'Animation plan must contain a slides array.'
    }

    $seenSlides = @{}
    foreach ($slidePlan in $Plan.slides) {
        $slideNumber = [int]$slidePlan.slide
        if ($slideNumber -lt 1) {
            throw "Invalid slide number: $slideNumber"
        }
        if ($seenSlides.ContainsKey($slideNumber)) {
            throw "Slide $slideNumber appears more than once in the animation plan."
        }
        $seenSlides[$slideNumber] = $true

        if ($null -eq $slidePlan.effects -or $slidePlan.effects.Count -lt 1) {
            throw "Slide $slideNumber has no effects."
        }
        if ($slidePlan.effects.Count -gt 8) {
            throw "Slide $slideNumber has more than 8 effects; reorganize or explicitly revise the plan."
        }

        $seenShapes = @{}
        foreach ($effect in $slidePlan.effects) {
            $shapeName = [string]$effect.shape
            if ([string]::IsNullOrWhiteSpace($shapeName)) {
                throw "Slide $slideNumber contains an effect without a shape name."
            }
            if ($seenShapes.ContainsKey($shapeName)) {
                throw "Slide $slideNumber repeats shape '$shapeName' in the main sequence."
            }
            $seenShapes[$shapeName] = $true

            if (-not $effectMap.ContainsKey([string]$effect.effect)) {
                throw "Unsupported effect '$($effect.effect)' on slide $slideNumber."
            }
            if (-not $triggerMap.ContainsKey([string]$effect.trigger)) {
                throw "Unsupported trigger '$($effect.trigger)' on slide $slideNumber."
            }
            if ($null -ne $effect.duration) {
                $duration = [double]$effect.duration
                if ($duration -lt 0.05 -or $duration -gt 2.0) {
                    throw "Duration for '$shapeName' on slide $slideNumber must be between 0.05 and 2.0 seconds."
                }
            }
        }
    }
}

function Test-PlanShapesInPptx {
    param(
        [string]$PptxPath,
        [object]$Plan
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
    try {
        foreach ($slidePlan in $Plan.slides) {
            $slideNumber = [int]$slidePlan.slide
            $entry = $zip.GetEntry("ppt/slides/slide$slideNumber.xml")
            if ($null -eq $entry) {
                throw "Animation plan refers to slide $slideNumber, but that slide XML does not exist."
            }

            $reader = [System.IO.StreamReader]::new($entry.Open())
            try {
                $xml = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }

            $shapeNames = @{}
            foreach ($match in [regex]::Matches($xml, '<p:cNvPr[^>]*\bname="([^"]+)"')) {
                $shapeNames[$match.Groups[1].Value] = $true
            }

            foreach ($plannedEffect in $slidePlan.effects) {
                $shape = $null
                $effect = $null
                $shapeName = [string]$plannedEffect.shape
                if (-not $shapeNames.ContainsKey($shapeName)) {
                    throw "Shape '$shapeName' was not found in slide $slideNumber OOXML. Use PptxGenJS objectName and regenerate the deck."
                }
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}

$inputFull = [System.IO.Path]::GetFullPath($InputPptx)
$planFull = [System.IO.Path]::GetFullPath($PlanPath)
$outputFull = [System.IO.Path]::GetFullPath($OutputPptx)

if (-not (Test-Path -LiteralPath $inputFull -PathType Leaf)) {
    throw "Input PPTX not found: $inputFull"
}
if (-not (Test-Path -LiteralPath $planFull -PathType Leaf)) {
    throw "Animation plan not found: $planFull"
}
if ($inputFull -eq $outputFull) {
    throw 'InputPptx and OutputPptx must be different paths.'
}
if ((Test-Path -LiteralPath $outputFull) -and -not $Overwrite) {
    throw "Output already exists. Use -Overwrite only after confirming the exact target: $outputFull"
}

$plan = Get-Content -LiteralPath $planFull -Raw -Encoding utf8 | ConvertFrom-Json
Validate-Plan $plan
Test-PlanShapesInPptx -PptxPath $inputFull -Plan $plan

if ($ValidateOnly) {
    [pscustomobject]@{
        status = 'valid'
        slides = $plan.slides.Count
        effects = @($plan.slides | ForEach-Object { $_.effects.Count } | Measure-Object -Sum).Sum
    } | ConvertTo-Json
    exit 0
}

$outputDir = Split-Path -Parent $outputFull
if (-not (Test-Path -LiteralPath $outputDir)) {
    [void](New-Item -ItemType Directory -Path $outputDir)
}

$ppt = $null
$pres = $null
$verify = $null
$written = @()

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($inputFull, -1, 0, 0)

    foreach ($slidePlan in $plan.slides) {
        $slideNumber = [int]$slidePlan.slide
        if ($slideNumber -gt $pres.Slides.Count) {
            throw "Animation plan refers to slide $slideNumber, but the deck has only $($pres.Slides.Count) slides."
        }

        $slide = $pres.Slides.Item($slideNumber)
        try {
            if ($slidePlan.clearExisting -eq $true) {
                while ($slide.TimeLine.MainSequence.Count -gt 0) {
                    $slide.TimeLine.MainSequence.Item(1).Delete()
                }
            }

            foreach ($plannedEffect in $slidePlan.effects) {
                $shapeName = [string]$plannedEffect.shape
                try {
                    $shape = $slide.Shapes.Item($shapeName)
                }
                catch {
                    throw "Shape '$shapeName' was not found on slide $slideNumber."
                }

                try {
                    $effectId = $effectMap[[string]$plannedEffect.effect]
                    $triggerId = $triggerMap[[string]$plannedEffect.trigger]
                    $effect = $slide.TimeLine.MainSequence.AddEffect($shape, $effectId, 0, $triggerId, -1)
                    if ($null -ne $plannedEffect.duration) {
                        $effect.Timing.Duration = [double]$plannedEffect.duration
                    }
                    if ($null -ne $plannedEffect.delay) {
                        $effect.Timing.TriggerDelayTime = [double]$plannedEffect.delay
                    }
                    $written += [pscustomobject]@{
                        slide = $slideNumber
                        shape = $shapeName
                        effect = [string]$plannedEffect.effect
                        trigger = [string]$plannedEffect.trigger
                    }
                }
                finally {
                    Release-ComObject $effect
                    Release-ComObject $shape
                }
            }
        }
        finally {
            Release-ComObject $slide
        }
    }

    if ((Test-Path -LiteralPath $outputFull) -and $Overwrite) {
        Remove-Item -LiteralPath $outputFull -Force
    }
    $pres.SaveAs($outputFull, 24)
    $pres.Close()
    Release-ComObject $pres
    $pres = $null

    $verify = $ppt.Presentations.Open($outputFull, -1, 0, 0)
    foreach ($slidePlan in $plan.slides) {
        $slideNumber = [int]$slidePlan.slide
        $actualCount = $verify.Slides.Item($slideNumber).TimeLine.MainSequence.Count
        $expectedCount = $slidePlan.effects.Count
        if ($actualCount -lt $expectedCount) {
            throw "Slide $slideNumber reopened with $actualCount effects; expected at least $expectedCount."
        }
    }

    $verify.Close()
    Release-ComObject $verify
    $verify = $null

    [pscustomobject]@{
        status = 'passed'
        input = $inputFull
        output = $outputFull
        animatedSlides = $plan.slides.Count
        effectsWritten = $written.Count
    } | ConvertTo-Json -Depth 4
    exit 0
}
catch {
    [pscustomobject]@{
        status = 'failed'
        error = $_.Exception.Message
        output = $outputFull
        effectsWrittenBeforeFailure = $written.Count
    } | ConvertTo-Json -Depth 4
    exit 1
}
finally {
    if ($null -ne $verify) {
        try { $verify.Close() } catch {}
    }
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
    }
    Release-ComObject $verify
    Release-ComObject $pres
    Release-ComObject $ppt
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
