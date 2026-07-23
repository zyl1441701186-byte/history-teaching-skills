param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [int[]]$TitleExemptSlides = @(),

    [string]$OutputJson = "",

    [switch]$ThrowOnFailure
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-Ngrams([string]$Text, [int]$Size = 2) {
    $clean = [regex]::Replace($Text, '\s|\d+\s*/\s*\d+|MEDIA-P\d+-\d+', '')
    $set = [System.Collections.Generic.HashSet[string]]::new()
    if ($clean.Length -lt $Size) {
        if ($clean.Length -gt 0) { [void]$set.Add($clean) }
        return $set
    }
    for ($i = 0; $i -le $clean.Length - $Size; $i++) {
        [void]$set.Add($clean.Substring($i, $Size))
    }
    return $set
}

function Get-Jaccard($A, $B) {
    if ($A.Count -eq 0 -and $B.Count -eq 0) { return 1.0 }
    $intersection = 0
    foreach ($item in $A) { if ($B.Contains($item)) { $intersection++ } }
    $union = $A.Count + $B.Count - $intersection
    if ($union -eq 0) { return 0.0 }
    return [math]::Round($intersection / $union, 3)
}

$resolved = (Resolve-Path -LiteralPath $PptxPath).Path
$archive = [System.IO.Compression.ZipFile]::OpenRead($resolved)

try {
    $slideEntries = @($archive.Entries | Where-Object {
        $_.FullName -match '^ppt/slides/slide(\d+)\.xml$'
    } | Sort-Object {
        [int]([regex]::Match($_.FullName, 'slide(\d+)').Groups[1].Value)
    })

    $slideCount = $slideEntries.Count
    $slides = @()
    $errors = @()
    $warnings = @()
    $previous = $null

    foreach ($entry in $slideEntries) {
        $slideNumber = [int]([regex]::Match($entry.FullName, 'slide(\d+)').Groups[1].Value)
        $reader = [System.IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }

        $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')

        $textRuns = @($xml.SelectNodes('//a:t', $ns) | ForEach-Object { $_.InnerText })
        $allText = ($textRuns -join ' ')
        $names = @($xml.SelectNodes('//p:cNvPr', $ns) | ForEach-Object { $_.name })
        $prefix = 'S{0:D2}' -f $slideNumber
        $isExempt = $TitleExemptSlides -contains $slideNumber
        $hasL1 = @($names | Where-Object { $_ -eq "${prefix}_L1Title" }).Count -gt 0
        $hasL2 = @($names | Where-Object { $_ -eq "${prefix}_L2Title" }).Count -gt 0
        $hasCurrentTask = @($names | Where-Object { $_ -eq "${prefix}_CurrentTask" }).Count -gt 0

        if (-not $isExempt -and -not $hasL1) {
            $errors += "Slide $slideNumber missing ${prefix}_L1Title"
        }
        if (-not $isExempt -and -not $hasL2) {
            $errors += "Slide $slideNumber missing ${prefix}_L2Title"
        }
        if (-not $isExempt -and -not $hasCurrentTask) {
            $warnings += "Slide $slideNumber missing ${prefix}_CurrentTask"
        }

        $residuePatterns = [ordered]@{
            ai_marker = 'AI\s*\u751F\u6210|AI\s*Generated'
            production_note = '\u7248\u5F0F\u7528\u9014|\u5236\u4F5C\u63D0\u793A|\u5185\u90E8\u68C0\u67E5|\u5EFA\u8BAE\u7528\u65F6|\u5FC5\u8BB2\u9875|\u53EF\u538B\u7F29\u9875|\u6559\u5E08\u5907\u8BFE\u9875|\u53EF\u9690\u85CF|AI\s*\u8BBE\u8BA1'
            empty_square = '\u25A1'
            bracket_placeholder = '\uFF08\s*\uFF09|\(\s*\)'
            underline_placeholder = '_{3,}|\uFF3F{3,}'
            note_prompt = '^\s*\u7B14\u8BB0\s*$'
        }
        $residues = @()
        foreach ($pair in $residuePatterns.GetEnumerator()) {
            if ($allText -match $pair.Value) {
                $residues += $pair.Key
                $errors += "Slide $slideNumber contains $($pair.Key) residue"
            }
        }

        $forbiddenFooterElements = @()
        if (@($names | Where-Object { $_ -match '(?i)Page(Number|No)|Footer(Page|Note)|NotePrompt' }).Count -gt 0) {
            $message = "Slide $slideNumber contains a forbidden footer note or page-number object"
            $forbiddenFooterElements += $message
            $errors += $message
        }
        foreach ($run in $textRuns) {
            if ($run -match '^\s*\u7B14\u8BB0\s*$') {
                $message = "Slide $slideNumber contains the forbidden lower-left note prompt"
                $forbiddenFooterElements += $message
                $errors += $message
            }
            if ($run -match '^\s*(\d+)\s*/\s*(\d+)\s*$') {
                $message = "Slide $slideNumber contains a forbidden visible page number: $run"
                $forbiddenFooterElements += $message
                $errors += $message
            }
        }

        $geometries = @($xml.SelectNodes('//p:sp/p:spPr/a:prstGeom', $ns) | ForEach-Object { $_.prst })
        $rectangles = @($geometries | Where-Object { $_ -in @('rect', 'roundRect') }).Count
        $shapeCount = $xml.SelectNodes('//p:sp', $ns).Count
        $pictures = $xml.SelectNodes('//p:pic', $ns).Count
        $connectors = $xml.SelectNodes('//p:cxnSp', $ns).Count
        $mediaObjects = @($names | Where-Object { $_ -match 'MEDIA' }).Count

        if ($shapeCount -gt 0 -and ($rectangles / $shapeCount) -gt 0.85 -and $rectangles -gt 10) {
            $warnings += "Slide $slideNumber is rectangle-heavy: $rectangles of $shapeCount shapes"
        }

        $grams = Get-Ngrams $allText
        $similarity = $null
        if ($null -ne $previous) {
            $similarity = Get-Jaccard $previous.Grams $grams
            if ($similarity -ge 0.82) {
                $warnings += "Slides $($previous.Slide) and $slideNumber have high text similarity: $similarity"
            }
        }

        $slides += [pscustomobject]@{
            slide = $slideNumber
            title_exempt = $isExempt
            has_l1 = $hasL1
            has_l2 = $hasL2
            has_current_task = $hasCurrentTask
            shapes = $shapeCount
            rectangles = $rectangles
            pictures = $pictures
            connectors = $connectors
            media_objects = $mediaObjects
            has_timing = $xml.SelectNodes('//p:timing', $ns).Count -gt 0
            residues = $residues
            forbidden_footer_elements = $forbiddenFooterElements
            previous_slide_similarity = $similarity
        }
        $previous = [pscustomobject]@{ Slide = $slideNumber; Grams = $grams }
    }

    $report = [pscustomobject]@{
        pptx = $resolved
        slide_count = $slideCount
        passed = $errors.Count -eq 0
        total_pictures = ($slides | Measure-Object -Property pictures -Sum).Sum
        total_media_objects = ($slides | Measure-Object -Property media_objects -Sum).Sum
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
        throw "History classroom structure validation failed with $($errors.Count) error(s)."
    }
}
finally {
    $archive.Dispose()
}
