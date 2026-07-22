[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$fullPath = [System.IO.Path]::GetFullPath($PptxPath)
if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
    throw "PPTX not found: $fullPath"
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
try {
    $slideEntries = @(
        $zip.Entries |
            Where-Object { $_.FullName -match '^ppt/slides/slide(\d+)\.xml$' } |
            Sort-Object { [int]([regex]::Match($_.FullName, 'slide(\d+)\.xml').Groups[1].Value) }
    )

    $timedSlides = @()
    foreach ($entry in $slideEntries) {
        $reader = [System.IO.StreamReader]::new($entry.Open())
        try {
            $xml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $number = [int]([regex]::Match($entry.FullName, 'slide(\d+)\.xml').Groups[1].Value)
        if ($xml -match '<p:timing(?:\s|>)') {
            $timedSlides += $number
        }
    }

    [pscustomobject]@{
        file = $fullPath
        slideCount = $slideEntries.Count
        slidesWithTiming = $timedSlides.Count
        timedSlides = $timedSlides
    } | ConvertTo-Json -Depth 5
}
finally {
    $zip.Dispose()
}
