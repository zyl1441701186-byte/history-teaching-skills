param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [ValidateSet('final', 'content-draft')]
    [string]$DeliveryMode = 'final',

    [string]$TeacherNotesPath = '',

    [string]$OutputJson = '',

    [switch]$ThrowOnFailure
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$resolvedPptx = (Resolve-Path -LiteralPath $PptxPath).Path
$archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedPptx)

try {
    $slideEntries = @($archive.Entries | Where-Object {
        $_.FullName -match '^ppt/slides/slide(\d+)\.xml$'
    } | Sort-Object {
        [int]([regex]::Match($_.FullName, 'slide(\d+)').Groups[1].Value)
    })

    $errors = @()
    $warnings = @()
    $slides = @()
    $courseStandardSlides = @()
    $visibleCourseStandardSlides = @()
    $totalBlips = 0
    $leafName = [System.IO.Path]::GetFileName($resolvedPptx)
    if ($DeliveryMode -eq 'final' -and $leafName -notmatch '^\u8BFE\u5802\u8BFE\u4EF6\.pptx$') {
        $errors += "Final mode requires the canonical classroom deck filename; found $leafName"
    }
    if ($DeliveryMode -eq 'content-draft' -and $leafName -notmatch '^\u8BFE\u5802\u8BFE\u4EF6-\u5185\u5BB9\u65BD\u5DE5\u7A3F\.pptx$') {
        $errors += "Content-draft mode requires the canonical content-draft filename; found $leafName"
    }

    foreach ($entry in $slideEntries) {
        $slideNumber = [int]([regex]::Match($entry.FullName, 'slide(\d+)').Groups[1].Value)
        $reader = [System.IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }

        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')

        $textRuns = @($xml.SelectNodes('//a:t', $ns) | ForEach-Object { $_.InnerText })
        $allText = ($textRuns -join ' ')
        $compactText = [regex]::Replace($allText, '\s', '')
        $names = @($xml.SelectNodes('//p:cNvPr', $ns) | ForEach-Object { $_.name })
        $isHidden = $xml.DocumentElement.GetAttribute('show') -eq '0'
        $blips = $xml.SelectNodes('//a:blip', $ns).Count
        $tables = $xml.SelectNodes('//a:tbl', $ns).Count
        $connectors = $xml.SelectNodes('//p:cxnSp', $ns).Count
        $totalBlips += $blips

        $isCourseStandard = $allText -match '\u8BFE\u6807\u4F9D\u636E|2025\u5E74\u4FEE\u8BA2\u7248|2020\u5E74\u4FEE\u8BA2\u7248|\u7248\u672C\u53D8\u5316\u5206\u6790'
        if ($isCourseStandard) {
            $courseStandardSlides += $slideNumber
            if (-not $isHidden) {
                $visibleCourseStandardSlides += $slideNumber
                $errors += "Slide $slideNumber is a visible curriculum-standard appendix slide"
            }
        }

        $internalMarkers = @()
        $internalPattern = '\u6559\u5E08\u5907\u8BFE\u9875|\u53EF\u9690\u85CF|AI\s*\u8BBE\u8BA1|AI\s*\u751F\u6210|\u5EFA\u8BAE\u7528\u65F6|\u5FC5\u8BB2\u9875|\u53EF\u538B\u7F29\u9875'
        if ($allText -match $internalPattern) {
            $internalMarkers += 'internal-production-text'
            $errors += "Slide $slideNumber contains internal production text"
        }

        $placeholderCount = @($textRuns | Where-Object {
            $_ -match '\u5360\u4F4D\u7B26|MEDIA-P\d+-\d+|\u5F85\u63D2\u5165'
        }).Count
        if ($placeholderCount -gt 0) {
            $warnings += "Slide $slideNumber contains teacher-replaceable media placeholders"
        }

        if (@($textRuns | Where-Object { $_ -match '^\s*\u8C22\u8C22\s*$' }).Count -gt 0) {
            $errors += "Slide $slideNumber is a generic thank-you slide"
        }

        $l1Text = ''
        foreach ($shape in @($xml.SelectNodes('//p:sp', $ns))) {
            $nameNode = $shape.SelectSingleNode('./p:nvSpPr/p:cNvPr', $ns)
            if ($null -ne $nameNode -and $nameNode.name -match '_L1Title$') {
                $l1Text = (@($shape.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
                break
            }
        }
        $bareSectionPattern = '^\s*\u73AF\u8282\s*(?:\d+|\u4E00|\u4E8C|\u4E09|\u56DB|\u4E94|\u516D|\u4E03|\u516B|\u4E5D|\u5341)+\s*$'
        if ($l1Text -match $bareSectionPattern) {
            $errors += "Slide $slideNumber uses a non-semantic L1 section title: $l1Text"
        }

        $roleCounts = [ordered]@{
            evidence = @($names | Where-Object { $_ -match '(?i)Evidence|Material|Knowledge|Source' }).Count
            task = @($names | Where-Object { $_ -match '(?i)Task|Question' }).Count
            process = @($names | Where-Object { $_ -match '(?i)Process|Scaffold|Observe|Compare|Reason' }).Count
            answer = @($names | Where-Object { $_ -match '(?i)Answer' }).Count
            conclusion = @($names | Where-Object { $_ -match '(?i)Conclusion' }).Count
        }
        $roleTotal = @(
            $roleCounts.GetEnumerator() | Where-Object { $_.Value -gt 0 }
        ).Count

        $effectiveShapeTexts = @()
        foreach ($shape in @($xml.SelectNodes('//p:sp', $ns))) {
            $nameNode = $shape.SelectSingleNode('./p:nvSpPr/p:cNvPr', $ns)
            $shapeName = if ($null -ne $nameNode) { [string]$nameNode.name } else { '' }
            $shapeText = (@($shape.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
            $isFrameText = $shapeName -match '(?i)_L[123]Title$|Header|Footer|Brand|PageNumber'
            $isMediaInstruction = $shapeName -match '(?i)MEDIA|Placeholder' -or $shapeText -match 'MEDIA-P\d+-\d+|\u5360\u4F4D\u7B26|\u5F85\u63D2\u5165|\u6559\u5E08\u8865\u5165'
            if (-not $isFrameText -and -not $isMediaInstruction) {
                $effectiveShapeTexts += $shapeText
            }
        }
        $effectiveText = [regex]::Replace(($effectiveShapeTexts -join ''), '\s', '')

        $isLikelyTeachingSlide = -not $isHidden -and -not $isCourseStandard -and $slideNumber -ne 1
        $hasStructuralEvidence = $tables -gt 0 -or $connectors -gt 0 -or $roleCounts.evidence -gt 0
        $isHollow = $isLikelyTeachingSlide -and $effectiveText.Length -lt 90 -and -not $hasStructuralEvidence -and $roleTotal -lt 3
        if ($isHollow -and $DeliveryMode -eq 'final') {
            $errors += "Slide $slideNumber is content-hollow; media placeholders cannot replace classroom content"
        }
        elseif ($isHollow) {
            $warnings += "Slide $slideNumber is content-hollow and remains a content draft"
        }

        if ($isLikelyTeachingSlide -and $roleTotal -lt 3 -and $DeliveryMode -eq 'final') {
            $errors += "Slide $slideNumber exposes fewer than three classroom-chain roles"
        }
        elseif ($isLikelyTeachingSlide -and $roleTotal -lt 3) {
            $warnings += "Slide $slideNumber has fewer than three named classroom-chain roles"
        }

        $slides += [pscustomobject]@{
            slide = $slideNumber
            hidden = $isHidden
            curriculum_standard = $isCourseStandard
            characters = $compactText.Length
            effective_characters = $effectiveText.Length
            embedded_visuals = $blips
            tables = $tables
            connectors = $connectors
            placeholders = $placeholderCount
            internal_markers = $internalMarkers
            l1 = $l1Text
            content_hollow = $isHollow
            classroom_roles = $roleCounts
            classroom_role_total = $roleTotal
        }
    }

    if ($courseStandardSlides.Count -gt 0) {
        $firstAppendix = ($courseStandardSlides | Measure-Object -Minimum).Minimum
        $nonAppendixAfter = @($slides | Where-Object {
            $_.slide -gt $firstAppendix -and -not $_.curriculum_standard
        })
        if ($nonAppendixAfter.Count -gt 0) {
            $errors += 'Curriculum-standard slides are not a contiguous appendix at the end'
        }
    }

    $teachingSlides = @($slides | Where-Object {
        -not $_.hidden -and -not $_.curriculum_standard -and $_.slide -ne 1
    })
    $effectiveCharacterCounts = @($teachingSlides | ForEach-Object { [int]$_.effective_characters } | Sort-Object)
    $medianEffectiveCharacters = 0
    if ($effectiveCharacterCounts.Count -gt 0) {
        $mid = [math]::Floor($effectiveCharacterCounts.Count / 2)
        if ($effectiveCharacterCounts.Count % 2 -eq 0) {
            $medianEffectiveCharacters = [math]::Round(
                ($effectiveCharacterCounts[$mid - 1] + $effectiveCharacterCounts[$mid]) / 2,
                1
            )
        }
        else {
            $medianEffectiveCharacters = $effectiveCharacterCounts[$mid]
        }
    }
    $lowDensitySlides = @($teachingSlides | Where-Object {
        $_.effective_characters -lt 120 -and $_.classroom_role_total -lt 4
    } | ForEach-Object { $_.slide })
    if ($DeliveryMode -eq 'final' -and $medianEffectiveCharacters -lt 150) {
        $errors += "Median effective classroom content is $medianEffectiveCharacters characters; expected at least 150"
    }
    if ($lowDensitySlides.Count -gt [math]::Ceiling($teachingSlides.Count * 0.35)) {
        $errors += "$($lowDensitySlides.Count) of $($teachingSlides.Count) teaching slides are low-density without a near-complete classroom chain"
    }

    for ($i = 0; $i -lt $teachingSlides.Count - 1; $i++) {
        $pair = @($teachingSlides[$i], $teachingSlides[$i + 1])
        if ($pair[1].slide -ne $pair[0].slide + 1) { continue }
        $coveredRoles = @()
        foreach ($roleName in @('evidence', 'task', 'process', 'answer', 'conclusion')) {
            if (($pair[0].classroom_roles[$roleName] + $pair[1].classroom_roles[$roleName]) -gt 0) {
                $coveredRoles += $roleName
            }
        }
        if ($DeliveryMode -eq 'final' -and $coveredRoles.Count -lt 4) {
            $errors += "Slides $($pair[0].slide)-$($pair[1].slide) expose fewer than four of the five classroom-chain roles"
        }
        elseif ($coveredRoles.Count -lt 5) {
            $warnings += "Slides $($pair[0].slide)-$($pair[1].slide) do not yet complete the full five-role classroom chain"
        }
    }

    $semanticSections = @($teachingSlides | Where-Object { $_.l1 } | Group-Object -Property l1)
    foreach ($section in $semanticSections) {
        $coveredRoles = @()
        foreach ($roleName in @('evidence', 'task', 'process', 'answer', 'conclusion')) {
            $sectionRoleCount = 0
            foreach ($slide in $section.Group) {
                $sectionRoleCount += $slide.classroom_roles[$roleName]
            }
            if ($sectionRoleCount -gt 0) { $coveredRoles += $roleName }
        }
        if ($DeliveryMode -eq 'final' -and $coveredRoles.Count -lt 5) {
            $errors += "Section '$($section.Name)' does not complete all five classroom-chain roles"
        }
    }

    if ($totalBlips -eq 0) {
        $warnings += 'Deck contains no embedded visual files; teacher-replaceable placeholders are allowed but must not be counted as actual visual evidence'
    }

    $notesPageCount = $null
    $notesPlaceholderCount = $null
    if ($TeacherNotesPath) {
        $resolvedNotes = (Resolve-Path -LiteralPath $TeacherNotesPath).Path
        $notesText = Get-Content -LiteralPath $resolvedNotes -Raw -Encoding utf8
        $pageMatch = [regex]::Match($notesText, '\u603B\u9875\u6570[^\d]*(\d+)')
        if ($pageMatch.Success) {
            $notesPageCount = [int]$pageMatch.Groups[1].Value
            if ($notesPageCount -ne $slideEntries.Count) {
                $errors += "Teacher notes declare $notesPageCount slides but PPTX contains $($slideEntries.Count)"
            }
        }
        else {
            $warnings += 'Teacher notes do not declare a parseable total slide count'
        }

        $placeholderMatch = [regex]::Match($notesText, '\u5916\u90E8\u5A92\u4F53\u5360\u4F4D\u7B26\u6570\u91CF[^\d]*(\d+)')
        if ($placeholderMatch.Success) {
            $notesPlaceholderCount = [int]$placeholderMatch.Groups[1].Value
        }
    }

    $report = [pscustomobject]@{
        pptx = $resolvedPptx
        sha256 = (Get-FileHash -LiteralPath $resolvedPptx -Algorithm SHA256).Hash
        delivery_mode = $DeliveryMode
        slide_count = $slideEntries.Count
        passed = $errors.Count -eq 0
        classroom_ready = $DeliveryMode -eq 'final' -and $errors.Count -eq 0
        curriculum_standard_slides = $courseStandardSlides
        visible_curriculum_standard_slides = $visibleCourseStandardSlides
        total_embedded_visuals = $totalBlips
        median_effective_characters = $medianEffectiveCharacters
        low_density_slides = $lowDensitySlides
        notes_page_count = $notesPageCount
        notes_placeholder_count = $notesPlaceholderCount
        errors = $errors
        warnings = $warnings
        slides = $slides
    }

    $json = $report | ConvertTo-Json -Depth 8
    if ($OutputJson) {
        $parent = Split-Path -Parent $OutputJson
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        [System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.UTF8Encoding]::new($false))
    }
    $json

    if ($ThrowOnFailure -and -not $report.passed) {
        throw "History classroom readiness validation failed with $($errors.Count) error(s)."
    }
}
finally {
    $archive.Dispose()
}
